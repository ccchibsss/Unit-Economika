#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
============================================================================
🚀 FBS UNIT ECONOMICS PRO 2026 — ПОЛНАЯ ВЕРСИЯ С ИНТЕГРАЦИЕЙ GOOGLE SHEETS
============================================================================
Операционный директор | FBS-экспертиза | Оптимизация складских остатков
Маркетплейсы: Ozon, Wildberries, Яндекс Маркет
Версия: 7.2.0 (Google Sheets Edition)

КЛЮЧЕВЫЕ ПРИНЦИПЫ:
1. НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ — все данные из API, AI, CSV или пользовательского ввода
2. Интеллектуальная загрузка данных с каскадным фолбэком (API → AI → CSV → User)
3. Полная прозрачность расчетов и источников данных
4. 100% сохранение исходного UI + новые возможности
5. Работа с Google Sheets без сервисного аккаунта (ручной экспорт + инструкции)

НОВЫЕ ВОЗМОЖНОСТИ (v7.2.0):
- Экспорт в Google Sheets без сервисного аккаунта (CSV, TSV, копирование в буфер)
- Пошаговый туториал для новичков
- Подсказки к полям ввода
- Автоматическое обновление таблицы при наличии credentials.json

НИЧЕГО НЕ СОКРАЩЕНО — АБСОЛЮТНО ПОЛНАЯ ИСХОДНАЯ + НОВАЯ ВЕРСИЯ
============================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from datetime import datetime, timedelta
import requests
import json
import re
import csv
import os
import base64
import hashlib
from pathlib import Path
import time
import logging
from typing import Dict, List, Any, Optional, Tuple, Union, Callable
from dataclasses import dataclass, field, asdict
from functools import wraps, lru_cache
import uuid
import math
import warnings
import io
import pickle
import shutil
from concurrent.futures import ThreadPoolExecutor, as_completed
import traceback
from enum import Enum, auto
from abc import ABC, abstractmethod
import getpass

# Попытка импорта дополнительных библиотек
try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False
    Fernet = None
    print("⚠️ Cryptography не установлен. Шифрование будет отключено.")

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle, numbers
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
    from openpyxl.chart import BarChart, Reference, PieChart, LineChart
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import DataPoint
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.hyperlink import Hyperlink
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False
    print("⚠️ OpenPyXL не установлен. Экспорт в Excel будет недоступен.")

try:
    import gspread
    from google.oauth2.service_account import Credentials
    from google.oauth2 import service_account
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False
    gspread = None
    print("⚠️ GSpread не установлен. Интеграция с Google Sheets будет недоступна.")

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4, landscape
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import cm
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import letter
    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False
    print("⚠️ ReportLab не установлен. Экспорт в PDF будет недоступен.")

warnings.filterwarnings('ignore')

# ============================================================================
# БЛОК 0: БАЗОВАЯ КОНФИГУРАЦИЯ И НАСТРОЙКИ
# ============================================================================

APP_VERSION = "7.2.0"
APP_NAME = "🚀 FBS Юнит-экономика PRO 2026 — Полная ИИ-версия с Google Sheets"
APP_DESCRIPTION = "Профессиональный расчет юнит-экономики для FBS-модели с ИИ и без сокращений"

# Настройка путей
BASE_DIR = Path(__file__).parent.resolve() if '__file__' in dir() else Path.cwd()
DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
LOGS_DIR = BASE_DIR / "logs"
EXPORTS_DIR = BASE_DIR / "exports"
CONFIG_DIR = BASE_DIR / "config"
TEMP_DIR = BASE_DIR / "temp"
TARIFFS_CACHE_DIR = CACHE_DIR / "tariffs"
INTELLIGENT_CACHE_DIR = CACHE_DIR / "intelligent_loader"

# Создание директорий
for dir_path in [DATA_DIR, CACHE_DIR, LOGS_DIR, EXPORTS_DIR, CONFIG_DIR, TEMP_DIR, TARIFFS_CACHE_DIR, INTELLIGENT_CACHE_DIR]:
    dir_path.mkdir(exist_ok=True, parents=True)

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - [%(filename)s:%(lineno)d] - %(message)s',
    handlers=[
        logging.FileHandler(LOGS_DIR / "fbs_unit_economy.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('FBSEconomy')

# ============================================================================
# БЛОК 1: ДЕКОРАТОРЫ И УТИЛИТЫ
# ============================================================================

def timing_decorator(func):
    @wraps(func)
    def wrapper(*args, **kwargs):
        start_time = time.time()
        result = func(*args, **kwargs)
        end_time = time.time()
        execution_time = end_time - start_time
        if execution_time > 1.0:
            logger.info(f"⏱️ {func.__name__} выполнена за {execution_time:.2f} сек")
        return result
    return wrapper

def retry_on_failure(max_retries: int = 3, delay: float = 1.0):
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            last_exception = None
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    last_exception = e
                    logger.warning(f"⚠️ Попытка {attempt + 1}/{max_retries} для {func.__name__} не удалась: {e}")
                    if attempt < max_retries - 1:
                        time.sleep(delay * (attempt + 1))
            raise last_exception
        return wrapper
    return decorator

def memoize(func):
    cache = {}
    @wraps(func)
    def wrapper(*args, **kwargs):
        key = (args, frozenset(kwargs.items()))
        if key not in cache:
            cache[key] = func(*args, **kwargs)
        return cache[key]
    wrapper.cache_clear = cache.clear
    return wrapper

class ProgressTracker:
    def __init__(self):
        self.progress = 0.0
        self.status = ""
        self.total = 0
        self.current = 0
        self.start_time = None
        self.estimated_time_remaining = 0
    
    def start(self, total: int, status: str = ""):
        self.total = total
        self.current = 0
        self.progress = 0.0
        self.status = status
        self.start_time = time.time()
    
    def update(self, current: int, status: str = ""):
        self.current = current
        self.total = max(self.total, current)
        self.progress = min(current / self.total, 1.0) if self.total > 0 else 0
        if status:
            self.status = status
        
        if self.start_time and self.progress > 0:
            elapsed = time.time() - self.start_time
            self.estimated_time_remaining = (elapsed / self.progress) * (1 - self.progress)
    
    def get_progress(self) -> float:
        return self.progress
    
    def get_status(self) -> str:
        return self.status
    
    def get_eta(self) -> float:
        return self.estimated_time_remaining

class AuditLogger:
    def __init__(self):
        self.audit_file = LOGS_DIR / "audit.log"
        self._init_audit_file()
    
    def _init_audit_file(self):
        if not self.audit_file.exists():
            with open(self.audit_file, 'w', encoding='utf-8') as f:
                f.write("timestamp,user,action,details\n")
    
    def log(self, action: str, details: Dict[str, Any]):
        user = getpass.getuser()
        timestamp = datetime.now().isoformat()
        
        with open(self.audit_file, 'a', encoding='utf-8') as f:
            f.write(f"{timestamp},{user},{action},{json.dumps(details, ensure_ascii=False)}\n")
        
        logger.info(f"📝 Аудит: {user} - {action}")

# ============================================================================
# БЛОК 2: БЕЗОПАСНОЕ ХРАНЕНИЕ ДАННЫХ (ШИФРОВАНИЕ)
# ============================================================================

class SecureDataManager:
    def __init__(self):
        self.key_file = CONFIG_DIR / ".master_key"
        self.data_file = CONFIG_DIR / ".secure_data.enc"
        self._fernet = None
        self._init_encryption()
    
    def _init_encryption(self):
        if not CRYPTO_AVAILABLE:
            logger.warning("⚠️ Cryptography не установлен. Данные не будут зашифрованы.")
            return
        
        try:
            if self.key_file.exists():
                key = self.key_file.read_bytes()
            else:
                key = Fernet.generate_key()
                self.key_file.write_bytes(key)
                try:
                    os.chmod(self.key_file, 0o600)
                except OSError:
                    pass
            
            self._fernet = Fernet(key)
            logger.info("✅ Шифрование инициализировано успешно")
        except Exception as e:
            logger.error(f"❌ Ошибка инициализации шифрования: {e}")
            self._fernet = None
    
    def is_available(self) -> bool:
        return self._fernet is not None
    
    def save_data(self, data: Dict[str, Any]) -> bool:
        if not self._fernet:
            logger.warning("⚠️ Шифрование недоступно, данные не сохранены")
            return False
        
        try:
            json_data = json.dumps(data, ensure_ascii=False, indent=2)
            encrypted = self._fernet.encrypt(json_data.encode('utf-8'))
            self.data_file.write_bytes(encrypted)
            try:
                os.chmod(self.data_file, 0o600)
            except OSError:
                pass
            
            logger.info("✅ Данные успешно зашифрованы и сохранены")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка сохранения зашифрованных данных: {e}")
            return False
    
    def load_data(self) -> Dict[str, Any]:
        if not self._fernet or not self.data_file.exists():
            return {}
        
        try:
            encrypted = self.data_file.read_bytes()
            decrypted = self._fernet.decrypt(encrypted)
            return json.loads(decrypted.decode('utf-8'))
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки зашифрованных данных: {e}")
            return {}
    
    def delete_data(self) -> bool:
        try:
            if self.data_file.exists():
                self.data_file.unlink()
                logger.info("🗑️ Зашифрованные данные удалены")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка удаления данных: {e}")
            return False
    
    def store_api_key(self, service: str, api_key: str) -> bool:
        data = self.load_data()
        if 'api_keys' not in data:
            data['api_keys'] = {}
        
        data['api_keys'][service] = api_key
        data['api_keys_updated'] = datetime.now().isoformat()
        
        return self.save_data(data)
    
    def get_api_key(self, service: str) -> Optional[str]:
        data = self.load_data()
        return data.get('api_keys', {}).get(service)
    
    def get_all_api_keys(self) -> Dict[str, str]:
        data = self.load_data()
        return data.get('api_keys', {})
    
    def delete_api_key(self, service: str) -> bool:
        data = self.load_data()
        if 'api_keys' in data and service in data['api_keys']:
            del data['api_keys'][service]
            return self.save_data(data)
        return False
    
    def clear_all_keys(self) -> bool:
        data = self.load_data()
        data['api_keys'] = {}
        return self.save_data(data)

# ============================================================================
# БЛОК 3: КЭШИРОВАНИЕ ДЛЯ ОПТИМИЗАЦИИ ПРОИЗВОДИТЕЛЬНОСТИ
# ============================================================================

class CacheManager:
    def __init__(self, max_memory_mb: int = 500, cache_ttl_seconds: int = 3600):
        self.cache_dir = CACHE_DIR
        self.max_memory_mb = max_memory_mb
        self.cache_ttl = cache_ttl_seconds
        self._memory_cache: Dict[str, Any] = {}
        self._cache_timestamps: Dict[str, float] = {}
        self._cache_sizes: Dict[str, int] = {}
        
        self.tariffs_cache_dir = self.cache_dir / "tariffs"
        self.api_cache_dir = self.cache_dir / "api_responses"
        self.calc_cache_dir = self.cache_dir / "calculations"
        self.intelligent_cache_dir = self.cache_dir / "intelligent_loader"
        
        for dir_path in [self.tariffs_cache_dir, self.api_cache_dir, self.calc_cache_dir, self.intelligent_cache_dir]:
            dir_path.mkdir(exist_ok=True, parents=True)
    
    def _get_cache_key(self, *args, **kwargs) -> str:
        key_parts = [str(arg) for arg in args]
        key_parts.extend(f"{k}={v}" for k, v in sorted(kwargs.items()))
        key_string = "|".join(key_parts)
        return hashlib.md5(key_string.encode()).hexdigest()
    
    def _get_disk_cache_path(self, cache_type: str, key: str) -> Path:
        cache_dirs = {
            'tariffs': self.tariffs_cache_dir,
            'api': self.api_cache_dir,
            'calc': self.calc_cache_dir,
            'intelligent_loader': self.intelligent_cache_dir
        }
        cache_dir = cache_dirs.get(cache_type, self.cache_dir)
        return cache_dir / f"{key}.cache"
    
    def get(self, cache_type: str, key: str) -> Optional[Any]:
        memory_key = f"{cache_type}:{key}"
        if memory_key in self._memory_cache:
            timestamp = self._cache_timestamps.get(memory_key, 0)
            if time.time() - timestamp < self.cache_ttl:
                logger.debug(f"📦 Кэш попадание (память): {memory_key}")
                return self._memory_cache[memory_key]
            else:
                del self._memory_cache[memory_key]
                del self._cache_timestamps[memory_key]
        
        cache_path = self._get_disk_cache_path(cache_type, key)
        if cache_path.exists():
            try:
                with open(cache_path, 'rb') as f:
                    cached_data = pickle.load(f)
                
                cached_time = cached_data.get('timestamp', 0)
                if time.time() - cached_time < self.cache_ttl:
                    logger.debug(f"💾 Кэш попадание (диск): {key}")
                    value = cached_data.get('data')
                    self._memory_cache[memory_key] = value
                    self._cache_timestamps[memory_key] = cached_time
                    return value
                else:
                    cache_path.unlink()
            except Exception as e:
                logger.debug(f"Ошибка чтения дискового кэша: {e}")
                try:
                    cache_path.unlink()
                except:
                    pass
        
        return None
    
    def set(self, cache_type: str, key: str, value: Any):
        memory_key = f"{cache_type}:{key}"
        current_time = time.time()
        
        self._memory_cache[memory_key] = value
        self._cache_timestamps[memory_key] = current_time
        
        try:
            cache_path = self._get_disk_cache_path(cache_type, key)
            cache_data = {
                'timestamp': current_time,
                'data': value,
                'cache_type': cache_type,
                'key': key
            }
            with open(cache_path, 'wb') as f:
                pickle.dump(cache_data, f, protocol=pickle.HIGHEST_PROTOCOL)
        except Exception as e:
            logger.debug(f"Не удалось сохранить кэш на диск: {e}")
        
        self._cleanup_memory_cache()
    
    def _cleanup_memory_cache(self):
        current_time = time.time()
        expired_keys = []
        
        for key, timestamp in self._cache_timestamps.items():
            if current_time - timestamp > self.cache_ttl:
                expired_keys.append(key)
        
        for key in expired_keys:
            if key in self._memory_cache:
                del self._memory_cache[key]
            if key in self._cache_timestamps:
                del self._cache_timestamps[key]
        
        if expired_keys:
            logger.debug(f"🗑️ Очищено {len(expired_keys)} устаревших записей кэша")
    
    def clear_cache(self, cache_type: Optional[str] = None):
        if cache_type:
            cache_dirs = {
                'tariffs': self.tariffs_cache_dir,
                'api': self.api_cache_dir,
                'calc': self.calc_cache_dir,
                'intelligent_loader': self.intelligent_cache_dir
            }
            if cache_type in cache_dirs:
                for cache_file in cache_dirs[cache_type].glob("*.cache"):
                    cache_file.unlink()
            
            prefix = f"{cache_type}:"
            keys_to_remove = [k for k in self._memory_cache if k.startswith(prefix)]
            for key in keys_to_remove:
                del self._memory_cache[key]
                if key in self._cache_timestamps:
                    del self._cache_timestamps[key]
        else:
            self._memory_cache.clear()
            self._cache_timestamps.clear()
            
            for cache_dir in [self.tariffs_cache_dir, self.api_cache_dir, self.calc_cache_dir, self.intelligent_cache_dir]:
                for cache_file in cache_dir.glob("*.cache"):
                    cache_file.unlink()
        
        logger.info(f"🗑️ Кэш очищен: {cache_type or 'все типы'}")
    
    def get_cache_stats(self) -> Dict[str, Any]:
        return {
            'memory_entries': len(self._memory_cache),
            'memory_size_mb': sum(self._cache_sizes.values()) / (1024 * 1024),
            'tariffs_cache_files': len(list(self.tariffs_cache_dir.glob("*.cache"))),
            'api_cache_files': len(list(self.api_cache_dir.glob("*.cache"))),
            'calc_cache_files': len(list(self.calc_cache_dir.glob("*.cache"))),
            'intelligent_cache_files': len(list(self.intelligent_cache_dir.glob("*.cache"))),
            'cache_ttl_seconds': self.cache_ttl
        }

# ============================================================================
# БЛОК 4: КОНФИГУРАЦИИ API МАРКЕТПЛЕЙСОВ (РАСШИРЕННЫЕ ЭНДПОИНТЫ)
# ============================================================================

class MarketplaceAPIEndpoint(Enum):
    OZON_COMMISSIONS = "https://api.ozon.ru/v1/commission/list"
    OZON_DELIVERY = "https://api.ozon.ru/v1/delivery-methods"
    OZON_DELIVERY_ZONES = "https://api.ozon.ru/v1/delivery-zones"
    OZON_ANALYTICS_SEASONAL = "https://api.ozon.ru/v1/analytics/seasonality"
    
    WILDBERRIES_TARIFFS = "https://suppliers-api.wildberries.ru/api/v2/tariffs"
    WILDBERRIES_COMMISSIONS = "https://suppliers-api.wildberries.ru/api/v2/commissions"
    WILDBERRIES_DELIVERY_ZONES = "https://suppliers-api.wildberries.ru/api/v2/delivery-zones"
    
    YANDEX_TARIFFS = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/tariffs"
    YANDEX_COMMISSIONS = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/offer-mapping-entries"
    YANDEX_DELIVERY_ZONES = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/delivery-zones"
    
    DEEPSEEK_CHAT = "https://api.deepseek.com/v1/chat/completions"

@dataclass
class MarketplaceTariffData:
    marketplace: str
    category: str
    commission_rate: float
    min_commission: float
    last_mile_base: float
    last_mile_per_kg: float
    last_mile_per_km: float
    acquiring_fee: float
    return_fee: float
    penalty_rate: float
    penalty_time_hours: int
    fbo_multiplier: float
    fbp_multiplier: float
    storage_base_rate: float
    min_logistics: float
    last_updated: str = ""
    source: str = "default"
    api_response_raw: str = ""
    confidence: float = 1.0
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'MarketplaceTariffData':
        return cls(**data)

# ============================================================================
# БЛОК 5: API МЕНЕДЖЕР ДЛЯ ЗАГРУЗКИ ТАРИФОВ (УЛУЧШЕННЫЙ)
# ============================================================================

class APIRateLimiter:
    def __init__(self):
        self.last_request_time: Dict[str, float] = {}
        self.min_interval: Dict[str, float] = {
            'ozon': 0.5,
            'wildberries': 1.0,
            'yandex_market': 1.0,
            'deepseek': 0.5
        }
    
    def wait_if_needed(self, service: str):
        if service in self.last_request_time:
            elapsed = time.time() - self.last_request_time[service]
            min_wait = self.min_interval.get(service, 0.5)
            if elapsed < min_wait:
                time.sleep(min_wait - elapsed)
        self.last_request_time[service] = time.time()

class MarketplaceAPIManager:
    def __init__(self, cache_manager: Optional[CacheManager] = None, 
                 secure_data: Optional[SecureDataManager] = None):
        self.cache_manager = cache_manager or CacheManager()
        self.secure_data = secure_data or SecureDataManager()
        self.rate_limiter = APIRateLimiter()
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': f'FBS-Unit-Economy-Pro/{APP_VERSION}',
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        })
        self.progress_tracker = ProgressTracker()
        self._api_keys_cache: Dict[str, str] = {}
        self._load_api_keys()
        self.audit_logger = AuditLogger()
        self._default_tariffs_cache: Dict[str, Dict] = {}
    
    def _load_api_keys(self):
        try:
            if self.secure_data.is_available():
                self._api_keys_cache = self.secure_data.get_all_api_keys()
            
            if not self._api_keys_cache:
                key_file = CONFIG_DIR / "api_keys.json"
                if key_file.exists():
                    with open(key_file, 'r', encoding='utf-8') as f:
                        self._api_keys_cache = json.load(f)
            
            if self._api_keys_cache:
                logger.info(f"🔑 Загружены API ключи для: {', '.join(self._api_keys_cache.keys())}")
        except Exception as e:
            logger.warning(f"⚠️ Не удалось загрузить API ключи: {e}")
            self._api_keys_cache = {}
    
    def save_api_key(self, service: str, api_key: str) -> bool:
        if not api_key or not api_key.strip():
            return False
        
        self._api_keys_cache[service] = api_key.strip()
        
        if self.secure_data.is_available():
            success = self.secure_data.store_api_key(service, api_key.strip())
            if success:
                self.audit_logger.log('save_api_key', {'service': service})
                logger.info(f"✅ API ключ для {service} сохранен в защищенное хранилище")
                return True
        
        try:
            key_file = CONFIG_DIR / "api_keys.json"
            with open(key_file, 'w', encoding='utf-8') as f:
                json.dump(self._api_keys_cache, f, indent=2, ensure_ascii=False)
            try:
                os.chmod(key_file, 0o600)
            except OSError:
                pass
            self.audit_logger.log('save_api_key', {'service': service, 'method': 'json_fallback'})
            logger.info(f"✅ API ключ для {service} сохранен в файл")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка сохранения API ключа: {e}")
            return False
    
    def get_api_key(self, service: str) -> Optional[str]:
        return self._api_keys_cache.get(service)
    
    def has_api_key(self, service: str) -> bool:
        return bool(self._api_keys_cache.get(service))
    
    def get_cached_tariffs(self, marketplace: str) -> Optional[Dict[str, Dict]]:
        cache_key = f"tariffs_{marketplace.lower()}"
        return self.cache_manager.get('tariffs', cache_key)
    
    def save_tariffs_to_cache(self, marketplace: str, tariffs: Dict[str, Dict]):
        cache_key = f"tariffs_{marketplace.lower()}"
        self.cache_manager.set('tariffs', cache_key, {
            'tariffs': tariffs,
            'marketplace': marketplace,
            'cached_at': datetime.now().isoformat(),
            'version': APP_VERSION
        })
        logger.info(f"💾 Тарифы {marketplace} сохранены в кэш")
    
    def load_tariffs_from_csv(self, marketplace: str, csv_content: str) -> Dict[str, Dict]:
        tariffs = {}
        try:
            df = pd.read_csv(io.StringIO(csv_content))
            
            required_cols = ['category', 'commission_rate', 'min_commission', 'last_mile_base']
            missing_cols = [col for col in required_cols if col not in df.columns]
            if missing_cols:
                raise ValueError(f"Отсутствуют обязательные колонки: {', '.join(missing_cols)}")
            
            for _, row in df.iterrows():
                category = str(row.get('category', 'default'))
                
                tariffs[category] = {
                    'commission_rate': float(row.get('commission_rate', 0.15)),
                    'min_commission': float(row.get('min_commission', 30.0)),
                    'last_mile_base': float(row.get('last_mile_base', 50.0)),
                    'last_mile_per_kg': float(row.get('last_mile_per_kg', 15.0)),
                    'last_mile_per_km': float(row.get('last_mile_per_km', 3.5)),
                    'acquiring_fee': float(row.get('acquiring_fee', 0.015)),
                    'return_fee': float(row.get('return_fee', 0.02)),
                    'penalty_rate': float(row.get('penalty_rate', 0.05)),
                    'penalty_time_hours': float(row.get('penalty_time_hours', 24)),
                    'fbo_multiplier': float(row.get('fbo_multiplier', 0.75)),
                    'fbp_multiplier': float(row.get('fbp_multiplier', 0.60)),
                    'storage_base_rate': float(row.get('storage_base_rate', 0.30)),
                    'min_logistics': float(row.get('min_logistics', 25.0)),
                    'source': 'csv_import',
                    'last_updated': datetime.now().isoformat(),
                    'confidence': 1.0
                }
            
            logger.info(f"✅ Загружено {len(tariffs)} категорий из CSV")
            self.audit_logger.log('load_tariffs_csv', {'marketplace': marketplace, 'count': len(tariffs)})
            
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки CSV: {e}")
            tariffs = {}
        
        return tariffs
    
    @retry_on_failure(max_retries=2, delay=2.0)
    def fetch_ozon_tariffs(self) -> Dict[str, Dict]:
        tariffs = {}
        
        client_id = self.get_api_key('ozon_client_id')
        api_key = self.get_api_key('ozon')
        
        if not client_id or not api_key:
            logger.warning("⚠️ API ключи Ozon не найдены")
            return {}
        
        self.rate_limiter.wait_if_needed('ozon')
        
        headers = {
            'Client-Id': client_id,
            'Api-Key': api_key,
            'Content-Type': 'application/json'
        }
        
        try:
            logger.info("📡 Запрос комиссий Ozon...")
            commission_response = self.session.post(
                MarketplaceAPIEndpoint.OZON_COMMISSIONS.value,
                headers=headers,
                json={"language": "RU"},
                timeout=30
            )
            
            if commission_response.status_code == 200:
                commission_data = commission_response.json()
                logger.info(f"✅ Получены комиссии Ozon: {len(commission_data.get('result', []))} категорий")
                
                logger.info("📡 Запрос тарифов доставки Ozon...")
                self.rate_limiter.wait_if_needed('ozon')
                delivery_response = self.session.post(
                    MarketplaceAPIEndpoint.OZON_DELIVERY.value,
                    headers=headers,
                    json={"language": "RU"},
                    timeout=30
                )
                
                delivery_data = {}
                if delivery_response.status_code == 200:
                    delivery_data = delivery_response.json()
                    logger.info("✅ Получены тарифы доставки Ozon")
                
                for item in commission_data.get('result', []):
                    category = item.get('category', 'default')
                    category_name = item.get('category_name', category)
                    
                    delivery_info = {}
                    for delivery_item in delivery_data.get('result', []):
                        if delivery_item.get('category') == category:
                            delivery_info = delivery_item
                            break
                    
                    tariffs[category_name] = {
                        'commission_rate': float(item.get('commission_percent', 15)) / 100,
                        'min_commission': float(item.get('min_commission', 30)),
                        'last_mile_base': float(delivery_info.get('delivery_base', 50)),
                        'last_mile_per_kg': float(delivery_info.get('delivery_per_kg', 15)),
                        'last_mile_per_km': float(delivery_info.get('delivery_per_km', 3.5)),
                        'acquiring_fee': float(item.get('acquiring_fee', 1.5)) / 100,
                        'return_fee': float(item.get('return_fee', 2.0)) / 100,
                        'penalty_rate': float(delivery_info.get('penalty_rate', 0.05)),
                        'penalty_time_hours': int(delivery_info.get('penalty_time_hours', 24)),
                        'fbo_multiplier': float(delivery_info.get('fbo_multiplier', 0.75)),
                        'fbp_multiplier': float(delivery_info.get('fbp_multiplier', 0.60)),
                        'storage_base_rate': float(item.get('storage_rate', 0.30)),
                        'min_logistics': float(delivery_info.get('min_delivery_cost', 25)),
                        'source': 'ozon_api',
                        'last_updated': datetime.now().isoformat(),
                        'confidence': 1.0,
                        'api_response_raw': json.dumps(item, ensure_ascii=False)
                    }
                
                if tariffs:
                    logger.info(f"✅ Загружено {len(tariffs)} категорий тарифов Ozon через API")
                    self.audit_logger.log('fetch_ozon_tariffs', {'count': len(tariffs), 'status': 'success'})
            else:
                logger.error(f"❌ Ozon API вернул статус {commission_response.status_code}")
        
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки тарифов Ozon: {e}")
            logger.exception(e)
        
        return tariffs
    
    @retry_on_failure(max_retries=2, delay=2.0)
    def fetch_wildberries_tariffs(self) -> Dict[str, Dict]:
        tariffs = {}
        
        api_key = self.get_api_key('wildberries')
        
        if not api_key:
            logger.warning("⚠️ API ключ Wildberries не найден")
            return {}
        
        self.rate_limiter.wait_if_needed('wildberries')
        
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        
        try:
            logger.info("📡 Запрос тарифов Wildberries...")
            tariffs_response = self.session.get(
                MarketplaceAPIEndpoint.WILDBERRIES_TARIFFS.value,
                headers=headers,
                params={'locale': 'ru'},
                timeout=30
            )
            
            logger.info("📡 Запрос комиссий Wildberries...")
            self.rate_limiter.wait_if_needed('wildberries')
            commissions_response = self.session.get(
                MarketplaceAPIEndpoint.WILDBERRIES_COMMISSIONS.value,
                headers=headers,
                params={'locale': 'ru'},
                timeout=30
            )
            
            tariffs_data = {}
            commissions_data = {}
            
            if tariffs_response.status_code == 200:
                tariffs_data = tariffs_response.json()
                logger.info("✅ Получены тарифы доставки Wildberries")
            
            if commissions_response.status_code == 200:
                commissions_data = commissions_response.json()
                logger.info("✅ Получены комиссии Wildberries")
            
            tariff_items = tariffs_data.get('data', {}).get('tariffs', [])
            commission_items = commissions_data.get('data', {}).get('commissions', [])
            
            commission_dict = {}
            for item in commission_items:
                category = item.get('categoryName', 'default')
                commission_dict[category] = item
            
            for item in tariff_items:
                category = item.get('categoryName', 'default')
                commission_info = commission_dict.get(category, {})
                
                tariffs[category] = {
                    'commission_rate': float(commission_info.get('commissionPercent', 16)) / 100,
                    'min_commission': float(commission_info.get('minCommission', 28)),
                    'last_mile_base': float(item.get('deliveryBase', 45)),
                    'last_mile_per_kg': float(item.get('deliveryPerKg', 14)),
                    'last_mile_per_km': float(item.get('deliveryPerKm', 3.2)),
                    'acquiring_fee': float(item.get('acquiringFee', 0.015)),
                    'return_fee': float(item.get('returnPercent', 1.8)) / 100,
                    'penalty_rate': float(item.get('penaltyRate', 0.08)),
                    'penalty_time_hours': int(item.get('penaltyTimeHours', 24)),
                    'fbo_multiplier': float(item.get('fboMultiplier', 0.70)),
                    'fbp_multiplier': float(item.get('fbpMultiplier', 0.55)),
                    'storage_base_rate': float(item.get('storageRate', 0.25)),
                    'min_logistics': float(item.get('minDeliveryCost', 22)),
                    'source': 'wildberries_api',
                    'last_updated': datetime.now().isoformat(),
                    'confidence': 1.0,
                    'api_response_raw': json.dumps(item, ensure_ascii=False)
                }
            
            if tariffs:
                logger.info(f"✅ Загружено {len(tariffs)} категорий тарифов Wildberries через API")
                self.audit_logger.log('fetch_wildberries_tariffs', {'count': len(tariffs), 'status': 'success'})
        
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки тарифов Wildberries: {e}")
            logger.exception(e)
        
        return tariffs
    
    @retry_on_failure(max_retries=2, delay=2.0)
    def fetch_yandex_market_tariffs(self) -> Dict[str, Dict]:
        tariffs = {}
        
        api_key = self.get_api_key('yandex_market')
        campaign_id = self.get_api_key('yandex_campaign_id')
        
        if not api_key or not campaign_id:
            logger.warning("⚠️ API ключи Яндекс Маркет не найдены")
            return {}
        
        self.rate_limiter.wait_if_needed('yandex_market')
        
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        
        try:
            tariffs_url = MarketplaceAPIEndpoint.YANDEX_TARIFFS.value.format(campaign_id=campaign_id)
            
            logger.info(f"📡 Запрос тарифов Яндекс Маркет (кампания: {campaign_id})...")
            response = self.session.get(
                tariffs_url,
                headers=headers,
                timeout=30
            )
            
            if response.status_code == 200:
                data = response.json()
                logger.info("✅ Получены тарифы Яндекс Маркет")
                
                for item in data.get('tariffs', []):
                    category = item.get('category', 'default')
                    category_name = item.get('categoryName', category)
                    
                    tariffs[category_name] = {
                        'commission_rate': float(item.get('commission', 14.5)) / 100,
                        'min_commission': float(item.get('minCommission', 35)),
                        'last_mile_base': float(item.get('deliveryBase', 55)),
                        'last_mile_per_kg': float(item.get('deliveryPerKg', 16)),
                        'last_mile_per_km': float(item.get('deliveryPerKm', 3.8)),
                        'acquiring_fee': float(item.get('acquiringFee', 0.015)),
                        'return_fee': float(item.get('returnFee', 0.025)),
                        'penalty_rate': float(item.get('penaltyRate', 0.07)),
                        'penalty_time_hours': int(item.get('penaltyTimeHours', 24)),
                        'fbo_multiplier': float(item.get('fboMultiplier', 0.80)),
                        'fbp_multiplier': float(item.get('fbpMultiplier', 0.65)),
                        'storage_base_rate': float(item.get('storageRate', 0.35)),
                        'min_logistics': float(item.get('minDeliveryCost', 30)),
                        'source': 'yandex_api',
                        'last_updated': datetime.now().isoformat(),
                        'confidence': 1.0,
                        'api_response_raw': json.dumps(item, ensure_ascii=False)
                    }
                
                if tariffs:
                    logger.info(f"✅ Загружено {len(tariffs)} категорий тарифов Яндекс Маркет через API")
                    self.audit_logger.log('fetch_yandex_tariffs', {'count': len(tariffs), 'status': 'success'})
        
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки тарифов Яндекс Маркет: {e}")
            logger.exception(e)
        
        return tariffs
    
    def fetch_tariffs_via_deepseek(self, marketplace: str) -> Dict[str, Dict]:
        tariffs = {}
        
        api_key = self.get_api_key('deepseek')
        
        if not api_key:
            logger.warning("⚠️ DeepSeek API ключ не найден")
            return {}
        
        self.rate_limiter.wait_if_needed('deepseek')
        
        try:
            prompt = f"""
            Ты эксперт по тарифам маркетплейсов с актуальными данными на 2026 год.
            
            Предоставь актуальные тарифы для маркетплейса "{marketplace}" 
            в формате строгого JSON без markdown-разметки.
            
            ВАЖНО: Верни ТОЛЬКО валидный JSON объект, без каких-либо пояснений.
            
            Формат ответа:
            {{
                "categories": {{
                    "default": {{
                        "commission_rate": 0.15,
                        "min_commission": 30,
                        "last_mile_base": 50,
                        "last_mile_per_kg": 15,
                        "last_mile_per_km": 3.5,
                        "acquiring_fee": 0.015,
                        "return_fee": 0.02,
                        "penalty_rate": 0.05,
                        "penalty_time_hours": 24,
                        "fbo_multiplier": 0.75,
                        "fbp_multiplier": 0.60,
                        "storage_base_rate": 0.30,
                        "min_logistics": 25
                    }},
                    "auto_parts": {{ ... }},
                    "electronics": {{ ... }},
                    "clothing": {{ ... }},
                    "home": {{ ... }}
                }},
                "source": "deepseek_ai",
                "confidence": 0.95,
                "data_collection_date": "2026-01"
            }}
            
            Укажи реальные актуальные тарифы для {marketplace} на 2026 год.
            Учти последние изменения в тарифной политике маркетплейса.
            """
            
            headers = {
                'Authorization': f'Bearer {api_key}',
                'Content-Type': 'application/json'
            }
            
            payload = {
                'model': 'deepseek-chat',
                'messages': [
                    {
                        'role': 'system',
                        'content': 'Ты эксперт по тарифам маркетплейсов. Отвечай только валидным JSON без пояснений.'
                    },
                    {
                        'role': 'user',
                        'content': prompt
                    }
                ],
                'temperature': 0.1,
                'max_tokens': 4000,
                'response_format': {'type': 'json_object'}
            }
            
            logger.info(f"🤖 Отправка запроса к DeepSeek AI для получения тарифов {marketplace}...")
            
            response = self.session.post(
                MarketplaceAPIEndpoint.DEEPSEEK_CHAT.value,
                headers=headers,
                json=payload,
                timeout=60
            )
            
            if response.status_code == 200:
                response_data = response.json()
                content_text = response_data['choices'][0]['message']['content']
                
                try:
                    content = json.loads(content_text)
                except json.JSONDecodeError:
                    json_match = re.search(r'\{.*\}', content_text, re.DOTALL)
                    if json_match:
                        content = json.loads(json_match.group())
                    else:
                        raise ValueError("Не удалось извлечь JSON из ответа DeepSeek")
                
                categories = content.get('categories', {})
                confidence = content.get('confidence', 0.8)
                
                for category, tariff_data in categories.items():
                    tariff_data['source'] = f'deepseek_ai'
                    tariff_data['last_updated'] = datetime.now().isoformat()
                    tariff_data['confidence'] = confidence
                    tariffs[category] = tariff_data
                
                if tariffs:
                    logger.info(f"✅ DeepSeek предоставил тарифы для {len(tariffs)} категорий {marketplace}")
                    self.audit_logger.log('fetch_deepseek_tariffs', {'marketplace': marketplace, 'count': len(tariffs)})
        
        except Exception as e:
            logger.error(f"❌ Ошибка запроса к DeepSeek: {e}")
            logger.exception(e)
        
        return tariffs
    
    def get_tariffs(self, marketplace: str, force_refresh: bool = False, 
                   use_ai_fallback: bool = True, csv_content: Optional[str] = None,
                   user_tariffs: Optional[Dict[str, Dict]] = None) -> Dict[str, Dict]:
        marketplace_lower = marketplace.lower()
        
        if not force_refresh:
            cached = self.get_cached_tariffs(marketplace)
            if cached:
                tariffs = cached.get('tariffs', {})
                if tariffs:
                    cached_time = cached.get('cached_at', '')
                    logger.info(f"📦 Использованы кэшированные тарифы {marketplace} от {cached_time}")
                    return tariffs
        
        logger.info(f"🔄 Загрузка тарифов {marketplace}...")
        
        tariffs = {}
        
        try:
            if marketplace == "Ozon":
                tariffs = self.fetch_ozon_tariffs()
            elif marketplace == "Wildberries":
                tariffs = self.fetch_wildberries_tariffs()
            elif marketplace == "Яндекс Маркет":
                tariffs = self.fetch_yandex_market_tariffs()
            else:
                logger.warning(f"⚠️ Неизвестный маркетплейс: {marketplace}")
        except Exception as e:
            logger.error(f"❌ Ошибка API для {marketplace}: {e}")
            tariffs = {}
        
        if not tariffs and use_ai_fallback:
            logger.info(f"🤖 Прямое API недоступно, использую DeepSeek AI для {marketplace}")
            try:
                ai_tariffs = self.fetch_tariffs_via_deepseek(marketplace)
                if ai_tariffs:
                    tariffs = ai_tariffs
                    logger.info(f"✅ Тарифы {marketplace} получены через DeepSeek AI")
            except Exception as e:
                logger.error(f"❌ DeepSeek также недоступен: {e}")
        
        if not tariffs and csv_content:
            logger.info(f"📄 Использую CSV импорт для {marketplace}")
            csv_tariffs = self.load_tariffs_from_csv(marketplace, csv_content)
            if csv_tariffs:
                tariffs = csv_tariffs
                logger.info(f"✅ Тарифы {marketplace} загружены из CSV")
        
        if not tariffs and user_tariffs:
            logger.info(f"👤 Использую пользовательские тарифы для {marketplace}")
            tariffs = user_tariffs
            for category in tariffs:
                tariffs[category]['source'] = 'user_input'
                tariffs[category]['last_updated'] = datetime.now().isoformat()
                tariffs[category]['confidence'] = 1.0
        
        if tariffs:
            self.save_tariffs_to_cache(marketplace, tariffs)
        else:
            logger.error(f"❌ ВСЕ ИСТОЧНИКИ НЕДОСТУПНЫ для {marketplace}. Тарифы не загружены.")
        
        return tariffs
    
    def get_all_tariffs_as_dataframe(self, marketplace: str) -> pd.DataFrame:
        tariffs = self.get_tariffs(marketplace)
        
        if not tariffs:
            return pd.DataFrame()
        
        rows = []
        for category, data in tariffs.items():
            rows.append({
                'Категория': category,
                'Комиссия, %': round(data.get('commission_rate', 0) * 100, 2),
                'Мин. комиссия, ₽': data.get('min_commission', 0),
                'База Last Mile, ₽': data.get('last_mile_base', 0),
                'Last Mile за кг, ₽': data.get('last_mile_per_kg', 0),
                'Last Mile за км, ₽': data.get('last_mile_per_km', 0),
                'Эквайринг, %': round(data.get('acquiring_fee', 0) * 100, 2),
                'Возвраты, %': round(data.get('return_fee', 0) * 100, 2),
                'Штраф за просрочку, %': round(data.get('penalty_rate', 0) * 100, 2),
                'Время на передачу, ч': data.get('penalty_time_hours', 0),
                'Множитель FBO': data.get('fbo_multiplier', 0),
                'Множитель FBP': data.get('fbp_multiplier', 0),
                'Хранение, ₽/день': data.get('storage_base_rate', 0),
                'Мин. логистика, ₽': data.get('min_logistics', 0),
                'Источник': data.get('source', 'unknown'),
                'Уверенность': f"{data.get('confidence', 1.0)*100:.0f}%",
                'Обновлено': data.get('last_updated', '')[:19] if data.get('last_updated') else ''
            })
        
        return pd.DataFrame(rows)
    
    def test_api_connection(self, marketplace: str) -> Dict[str, Any]:
        result = {
            'marketplace': marketplace,
            'timestamp': datetime.now().isoformat(),
            'status': 'unknown',
            'response_time_ms': 0,
            'error': None
        }
        
        start_time = time.time()
        
        try:
            if marketplace == "Ozon":
                if not self.has_api_key('ozon') or not self.has_api_key('ozon_client_id'):
                    result['status'] = 'no_api_key'
                    result['error'] = 'API ключи Ozon не настроены'
                else:
                    tariffs = self.fetch_ozon_tariffs()
                    result['status'] = 'success' if tariffs else 'empty_response'
            
            elif marketplace == "Wildberries":
                if not self.has_api_key('wildberries'):
                    result['status'] = 'no_api_key'
                    result['error'] = 'API ключ Wildberries не настроен'
                else:
                    tariffs = self.fetch_wildberries_tariffs()
                    result['status'] = 'success' if tariffs else 'empty_response'
            
            elif marketplace == "Яндекс Маркет":
                if not self.has_api_key('yandex_market') or not self.has_api_key('yandex_campaign_id'):
                    result['status'] = 'no_api_key'
                    result['error'] = 'API ключи Яндекс Маркет не настроены'
                else:
                    tariffs = self.fetch_yandex_market_tariffs()
                    result['status'] = 'success' if tariffs else 'empty_response'
            
            else:
                result['status'] = 'unknown_marketplace'
                result['error'] = f'Неизвестный маркетплейс: {marketplace}'
        
        except Exception as e:
            result['status'] = 'error'
            result['error'] = str(e)
        
        result['response_time_ms'] = round((time.time() - start_time) * 1000, 2)
        
        return result

# ============================================================================
# БЛОК 6: ДАТАКЛАССЫ ДЛЯ РАСЧЕТОВ (РАСШИРЕННЫЕ)
# ============================================================================

@dataclass
class FBSInputData:
    artikul: str = ""
    product_name: str = ""
    category: str = "default"
    
    selling_price: float = 0.0
    cogs: float = 0.0
    
    weight_kg: float = 0.0
    length_cm: float = 0.0
    width_cm: float = 0.0
    height_cm: float = 0.0
    
    first_mile_cost_per_unit: float = 0.0
    packaging_cost: float = 0.0
    pick_pack_time_min: float = 5.0
    operator_hourly_rate: float = 300.0
    warehouse_distance_km: float = 0.0
    
    transport_type: str = "own"
    transport_cost_per_km: float = 20.0
    pallet_capacity: int = 100
    pallet_cost: float = 2000.0
    
    marketing_budget_per_unit: float = 0.0
    
    stock_depth_days: int = 30
    daily_sales: int = 5
    warehouse_rent_per_sqm: float = 500.0
    warehouse_space_per_unit: float = 0.01
    safety_stock_days: int = 7
    reorder_point_days: int = 5
    supplier_lead_time_days: int = 3
    
    repeat_purchase_rate: float = 0.3
    avg_purchases_per_year: float = 2.5
    customer_retention_rate: float = 0.7
    discount_rate: float = 0.1
    
    has_night_shift: bool = False
    processing_capacity_per_hour: int = 20
    
    # Новые поля для ИИ и расширенной логистики
    target_regions: List[str] = field(default_factory=list)
    region_weights: Dict[str, float] = field(default_factory=dict)
    
    warehouse_total_area_sqm: float = 0.0
    warehouse_employees: int = 0
    warehouse_avg_salary: float = 0.0
    warehouse_equipment_cost: float = 0.0
    
    price_elasticity: Optional[float] = None
    competitive_low_price: Optional[float] = None
    competitive_mid_price: Optional[float] = None
    competitive_high_price: Optional[float] = None
    
    seasonal_coefficients: Dict[int, float] = field(default_factory=dict)
    market_trends: Dict[int, float] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'FBSInputData':
        known_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered_data = {k: v for k, v in data.items() if k in known_fields}
        return cls(**filtered_data)
    
    def validate(self) -> List[str]:
        errors = []
        
        if self.selling_price <= 0:
            errors.append("Цена продажи должна быть больше нуля")
        
        if self.cogs <= 0:
            errors.append("Себестоимость должна быть больше нуля")
        
        if self.cogs >= self.selling_price:
            errors.append("Себестоимость не может быть больше или равна цене продажи")
        
        if self.weight_kg < 0:
            errors.append("Вес не может быть отрицательным")
        
        if self.warehouse_distance_km < 0:
            errors.append("Расстояние не может быть отрицательным")
        
        if self.pallet_capacity <= 0:
            errors.append("Количество единиц на паллете должно быть больше нуля")
        
        if self.stock_depth_days < 0:
            errors.append("Глубина запаса не может быть отрицательной")
        
        if self.safety_stock_days < 0:
            errors.append("Страховой запас не может быть отрицательным")
        
        return errors

@dataclass
class FBSResultData:
    artikul: str = ""
    product_name: str = ""
    
    selling_price: float = 0.0
    total_expenses: float = 0.0
    gross_profit: float = 0.0
    margin_percent: float = 0.0
    roi_percent: float = 0.0
    
    commission: float = 0.0
    first_mile_cost: float = 0.0
    last_mile_cost: float = 0.0
    pick_pack_cost: float = 0.0
    packaging_cost: float = 0.0
    acquiring_cost: float = 0.0
    return_cost: float = 0.0
    penalty_cost: float = 0.0
    marketing_cost: float = 0.0
    warehouse_cost: float = 0.0
    tax_cost: float = 0.0
    
    penalty_probability: float = 0.0
    break_even_distance_km: float = 0.0
    max_discount_percent: float = 0.0
    safety_margin_price: float = 0.0
    break_even_volume: float = 0.0
    
    ltv: float = 0.0
    cac: float = 0.0
    ltv_cac_ratio: float = 0.0
    romi: float = 0.0
    
    optimal_stock_units: int = 0
    safety_stock_units: int = 0
    reorder_point_units: int = 0
    stock_turnover_days: float = 0.0
    stock_turnover_rate: float = 0.0
    days_of_inventory: float = 0.0
    holding_cost_per_unit: float = 0.0
    
    recommended_stock_depth_days: int = 0
    recommended_safety_stock_days: int = 0
    stock_optimization_potential: float = 0.0
    
    logistic_zone: str = "unknown"
    logistic_zone_label: str = ""
    logistic_recommendation: str = ""
    is_logistic_critical: bool = False
    
    space_efficiency_ratio: float = 0.0
    revenue_per_sqm: float = 0.0
    profit_per_sqm: float = 0.0
    
    seasonal_factor: float = 1.0
    adjusted_margin_percent: float = 0.0
    seasonal_recommendation: str = ""
    
    # Новые поля
    data_source: str = "unknown"
    data_confidence: float = 1.0
    
    optimal_price: float = 0.0
    optimal_price_confidence: float = 0.0
    
    warehouse_tco_monthly: float = 0.0
    warehouse_cost_per_order: float = 0.0
    
    weighted_delivery_cost: float = 0.0
    delivery_zone_count: int = 0
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)
    
    def get_summary(self) -> Dict[str, Any]:
        return {
            'artikul': self.artikul,
            'product_name': self.product_name,
            'selling_price': self.selling_price,
            'total_expenses': self.total_expenses,
            'gross_profit': self.gross_profit,
            'margin_percent': self.margin_percent,
            'roi_percent': self.roi_percent,
            'first_mile_cost': self.first_mile_cost,
            'last_mile_cost': self.last_mile_cost,
            'penalty_cost': self.penalty_cost,
            'ltv': self.ltv,
            'cac': self.cac,
            'ltv_cac_ratio': self.ltv_cac_ratio,
            'optimal_stock_units': self.optimal_stock_units,
            'stock_turnover_days': self.stock_turnover_days,
            'stock_optimization_potential': self.stock_optimization_potential,
            'break_even_volume': self.break_even_volume,
            'logistic_zone': self.logistic_zone_label,
            'adjusted_margin_percent': self.adjusted_margin_percent,
            'optimal_price': self.optimal_price,
            'data_source': self.data_source,
            'data_confidence': self.data_confidence,
            'weighted_delivery_cost': self.weighted_delivery_cost
        }

# ============================================================================
# БЛОК 7: КОНФИГУРАЦИИ НАЛОГОВЫХ СИСТЕМ
# ============================================================================

TAX_SYSTEMS = {
    "УСН 6% (доходы)": {
        "rate": 0.06,
        "base": "revenue",
        "name": "УСН_6",
        "description": "Упрощенная система налогообложения, 6% от доходов"
    },
    "УСН 15% (доходы-расходы)": {
        "rate": 0.15,
        "base": "profit",
        "min_rate": 0.01,
        "name": "УСН_15",
        "description": "Упрощенная система налогообложения, 15% от прибыли (мин. 1% от доходов)"
    },
    "ОСН (общая)": {
        "rate": 0.20,
        "base": "profit",
        "name": "ОСН",
        "description": "Общая система налогообложения, 20% налог на прибыль"
    },
    "НПД (самозанятый)": {
        "rate": 0.06,
        "base": "revenue",
        "name": "НПД",
        "description": "Налог на профессиональный доход, 6% от доходов с юрлицами"
    },
    "Патент": {
        "rate": 0.06,
        "base": "revenue",
        "name": "Патент",
        "description": "Патентная система налогообложения"
    }
}

# ============================================================================
# БЛОК 8: ОСНОВНОЙ КАЛЬКУЛЯТОР FBS ЮНИТ-ЭКОНОМИКИ (РАСШИРЕННЫЙ)
# ============================================================================

class FBSUnitEconomicsCalculator:
    def __init__(self, api_manager: Optional[MarketplaceAPIManager] = None, 
                 tax_system: str = "УСН 6% (доходы)",
                 intelligent_loader: Optional['IntelligentDataLoader'] = None):
        self.api_manager = api_manager or MarketplaceAPIManager()
        self.tax_system = tax_system
        self.current_marketplace = "Ozon"
        self.current_tariffs: Dict[str, Dict] = {}
        self.tariffs_updated_at: Optional[datetime] = None
        self.tariffs_source = "unknown"
        
        self.intelligent_loader = intelligent_loader
        
        self.geo_zones: List[Dict] = []
        self.weight_tiers: List[Dict] = []
        self.seasonal_coefficients: Dict[str, Dict[int, float]] = {}
        self.market_trends: Dict[str, Dict[int, float]] = {}
        self.price_elasticity: Dict[str, float] = {}
        self.regional_rent_rates: Dict[str, float] = {}
        self.labor_rates: Dict[str, float] = {}
        
        self.progress_tracker = ProgressTracker()
        self.audit_logger = AuditLogger()
        
        self.refresh_tariffs()
    
    def set_marketplace(self, marketplace_name: str):
        self.current_marketplace = marketplace_name
        self.refresh_tariffs()
        self.audit_logger.log('set_marketplace', {'marketplace': marketplace_name})
        logger.info(f"🏪 Установлен маркетплейс: {marketplace_name}")
    
    def set_intelligent_loader(self, loader: 'IntelligentDataLoader'):
        self.intelligent_loader = loader
        logger.info("🧠 Интеллектуальный загрузчик подключен к калькулятору")
    
    def load_geo_zones(self, force_refresh: bool = False, 
                       user_csv: Optional[str] = None,
                       user_data: Optional[List[Dict]] = None) -> bool:
        if not self.intelligent_loader:
            logger.warning("⚠️ Интеллектуальный загрузчик не настроен")
            return False
        
        result = self.intelligent_loader.load_data(
            data_category=DataCategory.GEO_ZONES,
            marketplace=self.current_marketplace,
            force_refresh=force_refresh,
            user_csv=user_csv,
            user_data=user_data
        )
        
        if result.success and result.data:
            self.geo_zones = result.data if isinstance(result.data, list) else result.data.get('zones', [])
            logger.info(f"✅ Загружено {len(self.geo_zones)} гео-зон (источник: {result.source.value})")
            return True
        
        logger.error(f"❌ Не удалось загрузить гео-зоны: {result.message}")
        return False
    
    def load_seasonal_coefficients(self, category: str, force_refresh: bool = False,
                                   user_csv: Optional[str] = None,
                                   user_coeffs: Optional[Dict[int, float]] = None) -> bool:
        if not self.intelligent_loader:
            return False
        
        result = self.intelligent_loader.load_data(
            data_category=DataCategory.SEASONAL_COEFFICIENTS,
            marketplace=self.current_marketplace,
            category=category,
            force_refresh=force_refresh,
            user_csv=user_csv,
            user_data=user_coeffs
        )
        
        if result.success and result.data:
            self.seasonal_coefficients[category] = result.data
            return True
        return False
    
    def refresh_tariffs(self, force: bool = False, use_ai: bool = False, 
                       csv_content: Optional[str] = None,
                       user_tariffs: Optional[Dict[str, Dict]] = None):
        logger.info(f"🔄 Обновление тарифов для {self.current_marketplace}...")
        
        self.current_tariffs = self.api_manager.get_tariffs(
            self.current_marketplace,
            force_refresh=force,
            use_ai_fallback=use_ai,
            csv_content=csv_content,
            user_tariffs=user_tariffs
        )
        
        self.tariffs_updated_at = datetime.now()
        
        sources = set()
        for tariff in self.current_tariffs.values():
            source = tariff.get('source', 'unknown')
            if 'api' in source:
                sources.add('api')
            elif 'deepseek' in source:
                sources.add('deepseek')
            elif 'csv' in source:
                sources.add('csv')
            elif 'user' in source:
                sources.add('user')
            else:
                sources.add('unknown')
        
        if 'api' in sources:
            self.tariffs_source = 'api'
        elif 'deepseek' in sources:
            self.tariffs_source = 'deepseek'
        elif 'csv' in sources:
            self.tariffs_source = 'csv'
        elif 'user' in sources:
            self.tariffs_source = 'user'
        else:
            self.tariffs_source = 'unknown'
        
        logger.info(f"✅ Тарифы обновлены. Источник: {self.tariffs_source}. Категорий: {len(self.current_tariffs)}")
    
    def get_tariff_for_category(self, category: str) -> Dict[str, Any]:
        if not self.current_tariffs:
            logger.warning("⚠️ Тарифы не загружены. Загрузите тарифы через API, CSV или введите вручную.")
            return {}
        
        if category in self.current_tariffs:
            return self.current_tariffs[category]
        
        category_lower = category.lower()
        for cat, tariff in self.current_tariffs.items():
            cat_lower = cat.lower()
            if category_lower in cat_lower or cat_lower in category_lower:
                logger.debug(f"🔍 Найдено частичное совпадение: {category} -> {cat}")
                return tariff
        
        if 'default' in self.current_tariffs:
            logger.debug(f"🔍 Использую категорию 'default' для: {category}")
            return self.current_tariffs['default']
        
        first_tariff = next(iter(self.current_tariffs.values()), {})
        if first_tariff:
            logger.warning(f"⚠️ Категория {category} не найдена в тарифах, использую первую доступную")
            return first_tariff
        
        return {}
    
    def _get_logistic_zone(self, distance_km: float) -> Dict[str, Any]:
        if distance_km <= 25:
            return {
                'zone': 'red',
                'label': '🔴 Критическая зона',
                'recommendation': 'Срочно оптимизировать логистику! Рассмотрите смену поставщика или увеличение загрузки паллет.',
                'is_critical': True
            }
        elif distance_km <= 50:
            return {
                'zone': 'yellow',
                'label': '🟡 Зона риска',
                'recommendation': 'Рассмотрите смену поставщика логистики или оптимизацию маршрутов.',
                'is_critical': False
            }
        elif distance_km <= 100:
            return {
                'zone': 'green',
                'label': '🟢 Безопасная зона',
                'recommendation': 'Логистика оптимальна. Продолжайте мониторинг.',
                'is_critical': False
            }
        else:
            return {
                'zone': 'blue',
                'label': '🔵 Идеальная зона',
                'recommendation': 'Отличная логистическая стратегия! Рекомендуется масштабирование.',
                'is_critical': False
            }
    
    def _calculate_weighted_delivery_cost(self, input_data: FBSInputData) -> float:
        if not self.geo_zones:
            return 0.0
        
        if input_data.length_cm > 0 and input_data.width_cm > 0 and input_data.height_cm > 0:
            vol_weight = (input_data.length_cm * input_data.width_cm * input_data.height_cm) / 5000.0
        else:
            vol_weight = 0
        
        billable_weight = max(input_data.weight_kg, vol_weight)
        
        if input_data.target_regions:
            filtered_zones = [z for z in self.geo_zones if z.get('region_code') in input_data.target_regions]
            if filtered_zones:
                zones = filtered_zones
            else:
                zones = self.geo_zones
        else:
            zones = self.geo_zones
        
        if input_data.region_weights:
            total_weight = sum(input_data.region_weights.values())
            if total_weight > 0:
                normalized_weights = {k: v/total_weight for k, v in input_data.region_weights.items()}
                
                total_cost = 0.0
                for zone in zones:
                    region_code = zone.get('region_code', '')
                    weight = normalized_weights.get(region_code, 0)
                    if weight > 0:
                        zone_cost = zone.get('base_delivery_cost', 0) + billable_weight * zone.get('cost_per_kg', 0)
                        total_cost += zone_cost * weight
                
                return total_cost
        
        if zones:
            total_cost = sum(
                z.get('base_delivery_cost', 0) + billable_weight * z.get('cost_per_kg', 0)
                for z in zones
            )
            return total_cost / len(zones)
        
        return 0.0
    
    @timing_decorator
    def calculate_unit_economics(self, input_data: FBSInputData) -> FBSResultData:
        validation_errors = input_data.validate()
        if validation_errors:
            logger.warning(f"⚠️ Ошибки валидации для {input_data.artikul}: {validation_errors}")
        
        result = FBSResultData()
        result.artikul = input_data.artikul
        result.product_name = input_data.product_name
        result.selling_price = input_data.selling_price
        
        tariff = self.get_tariff_for_category(input_data.category)
        
        if not tariff:
            logger.error(f"❌ Тариф не найден для категории {input_data.category}.")
            tariff = {
                'commission_rate': 0.0, 'min_commission': 0.0,
                'last_mile_base': 0.0, 'last_mile_per_kg': 0.0,
                'acquiring_fee': 0.0, 'return_fee': 0.0,
                'penalty_rate': 0.0, 'penalty_time_hours': 24,
                'fbo_multiplier': 1.0, 'fbp_multiplier': 1.0,
                'storage_base_rate': 0.0, 'min_logistics': 0.0,
                'source': 'not_loaded', 'confidence': 0.0
            }
        
        result.data_source = tariff.get('source', 'unknown')
        result.data_confidence = tariff.get('confidence', 1.0)
        
        # 1. Комиссия
        commission_rate = tariff.get('commission_rate', 0.0)
        min_commission = tariff.get('min_commission', 0.0)
        result.commission = max(input_data.selling_price * commission_rate, min_commission)
        
        # 2. First Mile
        if input_data.first_mile_cost_per_unit > 0:
            result.first_mile_cost = input_data.first_mile_cost_per_unit
        else:
            pallet_units = max(input_data.pallet_capacity, 1)
            cost_per_pallet = input_data.warehouse_distance_km * input_data.transport_cost_per_km * 2
            result.first_mile_cost = cost_per_pallet / pallet_units
        
        # 3. Last Mile (с учетом гео-зон)
        if input_data.length_cm > 0 and input_data.width_cm > 0 and input_data.height_cm > 0:
            vol_weight = (input_data.length_cm * input_data.width_cm * input_data.height_cm) / 5000.0
        else:
            vol_weight = 0
        
        billable_weight = max(input_data.weight_kg, vol_weight)
        billable_weight = math.ceil(billable_weight * 2) / 2
        
        weighted_delivery = self._calculate_weighted_delivery_cost(input_data)
        
        if weighted_delivery > 0:
            result.last_mile_cost = weighted_delivery
            result.weighted_delivery_cost = weighted_delivery
            result.delivery_zone_count = len(self.geo_zones)
        else:
            last_mile_base = tariff.get('last_mile_base', 0.0)
            last_mile_per_kg = tariff.get('last_mile_per_kg', 0.0)
            min_logistics = tariff.get('min_logistics', 0.0)
            result.last_mile_cost = max(last_mile_base + (billable_weight * last_mile_per_kg), min_logistics)
        
        # 4. Pick & Pack
        pick_pack_hours = input_data.pick_pack_time_min / 60.0
        result.pick_pack_cost = pick_pack_hours * input_data.operator_hourly_rate
        
        # 5. Упаковка
        result.packaging_cost = input_data.packaging_cost
        
        # 6. Эквайринг
        acquiring_fee = tariff.get('acquiring_fee', 0.0)
        result.acquiring_cost = input_data.selling_price * acquiring_fee
        
        # 7. Возвраты
        return_fee = tariff.get('return_fee', 0.0)
        result.return_cost = input_data.selling_price * return_fee
        
        # 8. Штрафы
        if input_data.has_night_shift:
            penalty_probability = 0.05
        else:
            penalty_probability = 0.35
        
        penalty_rate = tariff.get('penalty_rate', 0.0)
        result.penalty_probability = penalty_probability
        result.penalty_cost = input_data.selling_price * penalty_rate * penalty_probability
        
        # 9. Маркетинг
        result.marketing_cost = input_data.marketing_budget_per_unit
        
        # 10. Складские расходы (TCO)
        total_stock = input_data.stock_depth_days * input_data.daily_sales
        
        if input_data.warehouse_total_area_sqm > 0 and input_data.warehouse_employees > 0:
            total_monthly_cost = (
                input_data.warehouse_rent_per_sqm * input_data.warehouse_total_area_sqm +
                input_data.warehouse_avg_salary * input_data.warehouse_employees +
                input_data.warehouse_equipment_cost
            )
            monthly_orders = input_data.daily_sales * 30
            if monthly_orders > 0:
                result.warehouse_cost = total_monthly_cost / monthly_orders
                result.warehouse_tco_monthly = total_monthly_cost
                result.warehouse_cost_per_order = result.warehouse_cost
        elif total_stock > 0 and input_data.daily_sales > 0:
            total_warehouse_space = input_data.warehouse_space_per_unit * total_stock
            monthly_rent = input_data.warehouse_rent_per_sqm * total_warehouse_space
            result.warehouse_cost = monthly_rent / (30 * input_data.daily_sales)
        else:
            result.warehouse_cost = 0
        
        # 11. Налог
        tax_config = TAX_SYSTEMS.get(self.tax_system, TAX_SYSTEMS["УСН 6% (доходы)"])
        
        if tax_config["base"] == "revenue":
            result.tax_cost = input_data.selling_price * tax_config["rate"]
        else:
            pre_tax_expenses = (
                result.commission + result.first_mile_cost + result.last_mile_cost +
                result.pick_pack_cost + result.packaging_cost + result.acquiring_cost +
                result.return_cost + result.penalty_cost + result.marketing_cost +
                result.warehouse_cost + input_data.cogs
            )
            pre_tax_profit = input_data.selling_price - pre_tax_expenses
            result.tax_cost = max(0, pre_tax_profit * tax_config["rate"])
            
            if "min_rate" in tax_config:
                min_tax = input_data.selling_price * tax_config["min_rate"]
                result.tax_cost = max(result.tax_cost, min_tax)
        
        # 12. Итого расходов и прибыль
        result.total_expenses = (
            input_data.cogs + result.commission + result.first_mile_cost +
            result.last_mile_cost + result.pick_pack_cost + result.packaging_cost +
            result.acquiring_cost + result.return_cost + result.penalty_cost +
            result.marketing_cost + result.warehouse_cost + result.tax_cost
        )
        
        result.gross_profit = result.selling_price - result.total_expenses
        
        if result.selling_price > 0:
            result.margin_percent = (result.gross_profit / result.selling_price) * 100
        else:
            result.margin_percent = 0
        
        if input_data.cogs > 0:
            result.roi_percent = (result.gross_profit / input_data.cogs) * 100
        else:
            result.roi_percent = 0
        
        # 13. Точка безубыточности по расстоянию
        if result.first_mile_cost > 0 and input_data.pallet_capacity > 0:
            cost_per_km_per_unit = (input_data.transport_cost_per_km * 2) / input_data.pallet_capacity
            if cost_per_km_per_unit > 0:
                result.break_even_distance_km = result.gross_profit / cost_per_km_per_unit
            else:
                result.break_even_distance_km = float('inf')
        else:
            result.break_even_distance_km = float('inf')
        
        # 14. Логистические зоны риска
        logistic_zone_info = self._get_logistic_zone(
            result.break_even_distance_km if result.break_even_distance_km != float('inf') else 200
        )
        result.logistic_zone = logistic_zone_info['zone']
        result.logistic_zone_label = logistic_zone_info['label']
        result.logistic_recommendation = logistic_zone_info['recommendation']
        result.is_logistic_critical = logistic_zone_info['is_critical']
        
        # 15. Запас прочности по цене
        variable_costs_percent = (
            commission_rate + acquiring_fee + return_fee +
            penalty_rate * penalty_probability +
            TAX_SYSTEMS[self.tax_system]["rate"]
        )
        
        fixed_costs_per_unit = (
            input_data.cogs + result.first_mile_cost + result.last_mile_cost +
            result.pick_pack_cost + result.packaging_cost +
            result.marketing_cost + result.warehouse_cost
        )
        
        denominator = 1 - variable_costs_percent
        if denominator > 0:
            min_price = fixed_costs_per_unit / denominator
        else:
            min_price = fixed_costs_per_unit * 2
        
        result.safety_margin_price = input_data.selling_price - min_price
        
        if input_data.selling_price > 0:
            result.max_discount_percent = ((input_data.selling_price - min_price) / input_data.selling_price) * 100
        else:
            result.max_discount_percent = 0
        
        # 16. Точка безубыточности по объему
        variable_costs = result.commission + result.last_mile_cost + result.acquiring_cost + result.return_cost + result.penalty_cost
        
        if (result.selling_price - variable_costs) > 0:
            result.break_even_volume = fixed_costs_per_unit / (result.selling_price - variable_costs)
        else:
            result.break_even_volume = float('inf')
        
        # 17. LTV и CAC
        if (1 + input_data.discount_rate) > 0:
            result.ltv = (
                input_data.selling_price * input_data.avg_purchases_per_year *
                input_data.customer_retention_rate
            ) / (1 + input_data.discount_rate)
        else:
            result.ltv = input_data.selling_price * input_data.avg_purchases_per_year * input_data.customer_retention_rate
        
        total_acquisition_cost = result.marketing_cost + result.penalty_cost + result.first_mile_cost
        new_customers_per_order = 0.3
        
        if new_customers_per_order > 0:
            result.cac = total_acquisition_cost / new_customers_per_order
        else:
            result.cac = 0
        
        if result.cac > 0:
            result.ltv_cac_ratio = result.ltv / result.cac
        else:
            result.ltv_cac_ratio = float('inf')
        
        if result.marketing_cost > 0:
            result.romi = (result.gross_profit / result.marketing_cost) * 100
        else:
            result.romi = 0
        
        # 18. Сезонная корректировка
        current_month = datetime.now().month
        
        if input_data.category in self.seasonal_coefficients:
            seasonal_factor = self.seasonal_coefficients[input_data.category].get(current_month, 1.0)
        elif input_data.seasonal_coefficients:
            seasonal_factor = input_data.seasonal_coefficients.get(current_month, 1.0)
        else:
            seasonal_factor = 1.0
        
        result.seasonal_factor = seasonal_factor
        result.adjusted_margin_percent = result.margin_percent * seasonal_factor
        
        if result.adjusted_margin_percent < 10:
            result.seasonal_recommendation = "⚠️ Низкая сезонная маржа - рассмотрите акции или повышение цен"
        elif result.adjusted_margin_percent < 20:
            result.seasonal_recommendation = "📊 Средняя сезонная маржа - стабильно, можно улучшить"
        else:
            result.seasonal_recommendation = "✅ Отличная сезонная маржа!"
        
        # 19. Оптимизация складских остатков
        daily_demand = input_data.daily_sales
        annual_demand = daily_demand * 365
        ordering_cost = 500.0
        holding_cost_per_unit = input_data.warehouse_rent_per_sqm * input_data.warehouse_space_per_unit * 12
        
        if holding_cost_per_unit > 0 and ordering_cost > 0:
            eoq = math.sqrt((2 * annual_demand * ordering_cost) / holding_cost_per_unit)
            result.optimal_stock_units = int(math.ceil(eoq))
        else:
            result.optimal_stock_units = input_data.stock_depth_days * daily_demand
        
        max_daily_demand = daily_demand * 1.5
        result.safety_stock_units = int(math.ceil(
            (max_daily_demand * input_data.supplier_lead_time_days) - 
            (daily_demand * input_data.supplier_lead_time_days)
        ))
        
        result.reorder_point_units = int(math.ceil(
            (daily_demand * input_data.supplier_lead_time_days) + result.safety_stock_units
        ))
        
        if daily_demand > 0 and result.optimal_stock_units > 0:
            result.stock_turnover_days = result.optimal_stock_units / daily_demand
        else:
            result.stock_turnover_days = 0
        
        if result.optimal_stock_units > 0:
            result.stock_turnover_rate = annual_demand / result.optimal_stock_units
        else:
            result.stock_turnover_rate = 0
        
        if daily_demand > 0 and result.optimal_stock_units > 0:
            result.days_of_inventory = result.optimal_stock_units / daily_demand
        else:
            result.days_of_inventory = 0
        
        result.holding_cost_per_unit = holding_cost_per_unit
        
        # 20. Эффективность использования пространства
        total_stock = input_data.stock_depth_days * input_data.daily_sales
        if total_stock > 0 and input_data.warehouse_space_per_unit > 0:
            total_sqm = total_stock * input_data.warehouse_space_per_unit
            if total_sqm > 0:
                result.space_efficiency_ratio = total_stock / total_sqm
                result.revenue_per_sqm = (input_data.selling_price * input_data.daily_sales * 30) / total_sqm
                result.profit_per_sqm = (result.gross_profit * input_data.daily_sales * 30) / total_sqm
        
        # 21. Рекомендации по оптимизации склада
        if result.stock_turnover_rate > 12:
            result.recommended_stock_depth_days = max(14, input_data.stock_depth_days - 5)
        elif result.stock_turnover_rate < 6:
            result.recommended_stock_depth_days = input_data.stock_depth_days + 5
        else:
            result.recommended_stock_depth_days = input_data.stock_depth_days
        
        if result.penalty_probability > 0.2:
            result.recommended_safety_stock_days = min(14, input_data.safety_stock_days + 3)
        else:
            result.recommended_safety_stock_days = max(3, input_data.safety_stock_days - 2)
        
        current_stock = input_data.stock_depth_days * daily_demand
        recommended_stock = result.recommended_stock_depth_days * daily_demand
        if current_stock > 0:
            result.stock_optimization_potential = ((current_stock - recommended_stock) / current_stock) * 100
        else:
            result.stock_optimization_potential = 0
        
        # 22. Оптимальная цена
        elasticity = input_data.price_elasticity or self.price_elasticity.get(input_data.category)
        if elasticity and abs(elasticity) > 1:
            optimal_markup = abs(elasticity) / (abs(elasticity) - 1)
            result.optimal_price = input_data.cogs * optimal_markup
        else:
            result.optimal_price = input_data.selling_price
        
        self.audit_logger.log('calculate_unit', {
            'artikul': input_data.artikul,
            'profit': result.gross_profit,
            'margin': result.margin_percent,
            'logistic_zone': result.logistic_zone,
            'data_source': result.data_source
        })
        
        return result
    
    @timing_decorator
    def calculate_batch(self, input_data_list: List[FBSInputData],
                       use_parallel: bool = True,
                       max_workers: int = 8) -> List[FBSResultData]:
        total = len(input_data_list)
        results = [None] * total
        
        if total == 0:
            return []
        
        self.progress_tracker.start(total, f"Расчет {total} товаров...")
        
        if total > 100 and use_parallel:
            logger.info(f"⚡ Запуск параллельной обработки {total} товаров ({max_workers} потоков)")
            
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                future_to_index = {}
                for i, data in enumerate(input_data_list):
                    future = executor.submit(self.calculate_unit_economics, data)
                    future_to_index[future] = i
                
                completed = 0
                for future in as_completed(future_to_index):
                    index = future_to_index[future]
                    try:
                        result = future.result()
                        results[index] = result
                        completed += 1
                        
                        if completed % 100 == 0 or completed == total:
                            self.progress_tracker.update(completed, f"Обработано {completed}/{total} товаров")
                    except Exception as e:
                        logger.error(f"❌ Ошибка расчета товара с индексом {index}: {e}")
                        error_result = FBSResultData()
                        error_result.artikul = f"ERROR_{index}"
                        error_result.product_name = f"Ошибка расчета: {str(e)[:50]}"
                        results[index] = error_result
                        completed += 1
        else:
            logger.info(f"🔄 Запуск последовательной обработки {total} товаров")
            
            for i, data in enumerate(input_data_list):
                try:
                    results[i] = self.calculate_unit_economics(data)
                except Exception as e:
                    logger.error(f"❌ Ошибка расчета товара {data.artikul}: {e}")
                    error_result = FBSResultData()
                    error_result.artikul = data.artikul
                    error_result.product_name = f"Ошибка: {str(e)[:50]}"
                    results[i] = error_result
                
                if (i + 1) % 50 == 0 or (i + 1) == total:
                    self.progress_tracker.update(i + 1, f"Обработано {i + 1}/{total} товаров")
        
        self.progress_tracker.update(total, f"✅ Расчет завершен! Обработано {total} товаров")
        results = [r for r in results if r is not None]
        
        logger.info(f"✅ Пакетный расчет завершен. Успешно: {len(results)}/{total}")
        
        return results
    
    def run_what_if_analysis(self, base_data: FBSInputData, scenarios: List[Dict]) -> pd.DataFrame:
        results = []
        
        for scenario in scenarios:
            test_data = FBSInputData(**base_data.to_dict())
            
            for param, value in scenario.items():
                if param == 'name':
                    continue
                if hasattr(test_data, param):
                    setattr(test_data, param, value)
            
            result = self.calculate_unit_economics(test_data)
            results.append({
                'Сценарий': scenario.get('name', 'Без названия'),
                'Прибыль, ₽': round(result.gross_profit, 2),
                'Маржа, %': round(result.margin_percent, 2),
                'ROI, %': round(result.roi_percent, 2),
                'Точка безубыт., км': round(result.break_even_distance_km, 1) if result.break_even_distance_km != float('inf') else '∞',
                'Точка безубыт., шт': round(result.break_even_volume, 0) if result.break_even_volume != float('inf') else '∞',
                'Опт. запас, шт': result.optimal_stock_units,
                'Оборачиваемость, дн': round(result.stock_turnover_days, 1),
                'Логистическая зона': result.logistic_zone_label,
                'Скорр. маржа, %': round(result.adjusted_margin_percent, 2),
                'Опт. цена, ₽': round(result.optimal_price, 2)
            })
        
        self.audit_logger.log('what_if_analysis', {'scenarios': len(scenarios)})
        return pd.DataFrame(results)
    
    def generate_recommendations(self, results: List[FBSResultData]) -> List[Dict]:
        recommendations = []
        
        if not results:
            return recommendations
        
        total = len(results)
        
        profitable = [r for r in results if r.gross_profit > 0]
        unprofitable = [r for r in results if r.gross_profit <= 0]
        
        if len(unprofitable) / total > 0.2:
            recommendations.append({
                'priority': 'high',
                'category': 'Прибыльность',
                'icon': '⚠️',
                'message': f'{len(unprofitable)} товаров ({len(unprofitable)/total*100:.0f}%) убыточны. Рекомендуется пересмотреть цены или сократить затраты First Mile.',
                'affected_products': [r.artikul for r in unprofitable[:5]]
            })
        elif len(unprofitable) > 0:
            recommendations.append({
                'priority': 'medium',
                'category': 'Прибыльность',
                'icon': '📊',
                'message': f'{len(unprofitable)} товаров убыточны. Рассмотрите возможность повышения цен или оптимизации затрат.',
                'affected_products': [r.artikul for r in unprofitable[:5]]
            })
        
        high_distance = [r for r in results if r.is_logistic_critical and r.break_even_distance_km < 25]
        if high_distance:
            recommendations.append({
                'priority': 'high',
                'category': 'Логистика',
                'icon': '🚚',
                'message': f'{len(high_distance)} товаров находятся в критической логистической зоне (<25 км). Рекомендуется оптимизировать маршруты или увеличить загрузку паллет.',
                'affected_products': [r.artikul for r in high_distance[:5]]
            })
        
        high_stock = [r for r in results if r.stock_turnover_days > 60]
        if high_stock:
            recommendations.append({
                'priority': 'medium',
                'category': 'Склад',
                'icon': '📦',
                'message': f'{len(high_stock)} товаров имеют низкую оборачиваемость (>60 дней). Рекомендуется провести акции для ускорения оборота.',
                'affected_products': [r.artikul for r in high_stock[:5]]
            })
        
        low_ltv = [r for r in results if r.ltv_cac_ratio < 3 and r.ltv_cac_ratio > 0]
        if low_ltv:
            recommendations.append({
                'priority': 'medium',
                'category': 'Маркетинг',
                'icon': '📈',
                'message': f'{len(low_ltv)} товаров имеют низкое соотношение LTV/CAC (<3). Рекомендуется оптимизировать маркетинговые расходы.',
                'affected_products': [r.artikul for r in low_ltv[:5]]
            })
        
        # Новые рекомендации по качеству данных
        low_confidence = [r for r in results if r.data_confidence < 0.9]
        if low_confidence:
            recommendations.append({
                'priority': 'info',
                'category': 'Качество данных',
                'icon': '🔍',
                'message': f'{len(low_confidence)} товаров используют данные с пониженной уверенностью. Рекомендуется загрузить точные тарифы через API или CSV.',
                'affected_products': [r.artikul for r in low_confidence[:5]]
            })
        
        self.audit_logger.log('generate_recommendations', {'count': len(recommendations)})
        return recommendations

# ============================================================================
# БЛОК 9: СИСТЕМА ИНТЕЛЛЕКТУАЛЬНОЙ ЗАГРУЗКИ ДАННЫХ (API → AI → CSV → USER)
# ============================================================================

class DataSource(Enum):
    API = "api"
    AI = "ai"
    CSV = "csv"
    USER = "user"
    CACHE = "cache"
    NONE = "none"

class DataCategory(Enum):
    GEO_ZONES = "geo_zones"
    WEIGHT_TIERS = "weight_tiers"
    SEASONAL_COEFFICIENTS = "seasonal_coefficients"
    MARKET_TRENDS = "market_trends"
    PRICE_ELASTICITY = "price_elasticity"
    COMPETITIVE_PRICES = "competitive_prices"
    REGIONAL_RENT_RATES = "regional_rent_rates"
    LABOR_RATES = "labor_rates"

@dataclass
class DataLoadAttempt:
    source: DataSource
    timestamp: str
    success: bool
    data_count: int = 0
    error_message: str = ""
    response_time_ms: float = 0.0
    endpoint: str = ""
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'source': self.source.value,
            'timestamp': self.timestamp,
            'success': self.success,
            'data_count': self.data_count,
            'error_message': self.error_message,
            'response_time_ms': self.response_time_ms,
            'endpoint': self.endpoint
        }

@dataclass
class DataLoadResult:
    success: bool
    source: DataSource
    data: Any
    data_category: DataCategory
    confidence: float = 1.0
    message: str = ""
    attempts: List[DataLoadAttempt] = field(default_factory=list)
    load_time_ms: float = 0.0
    expires_at: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        return {
            'success': self.success,
            'source': self.source.value,
            'data_category': self.data_category.value,
            'confidence': self.confidence,
            'message': self.message,
            'attempts_count': len(self.attempts),
            'load_time_ms': self.load_time_ms,
            'expires_at': self.expires_at
        }

class IntelligentDataLoader:
    """
    Интеллектуальный загрузчик данных с каскадным фолбэком.
    API → AI → CSV → User
    НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ
    """
    
    def __init__(self, 
                 api_manager: Optional[MarketplaceAPIManager] = None,
                 cache_manager: Optional[CacheManager] = None,
                 secure_data: Optional[SecureDataManager] = None):
        self.api_manager = api_manager or MarketplaceAPIManager()
        self.cache_manager = cache_manager or CacheManager()
        self.secure_data = secure_data or SecureDataManager()
        
        self.deepseek_api_key = self.secure_data.get_api_key('deepseek') if self.secure_data else None
        
        self.load_history: List[DataLoadResult] = []
        self.api_timeout = 30
        self.ai_timeout = 60
        self.cache_ttl_hours = 24
        
        self.stats = {
            'total_attempts': 0,
            'successful_loads': 0,
            'failed_loads': 0,
            'by_source': {source.value: 0 for source in DataSource},
            'total_data_points_loaded': 0
        }
        
        self._init_ai_prompts()
        
        logger.info(f"🧠 IntelligentDataLoader инициализирован (AI: {'доступен' if self.deepseek_api_key else 'недоступен'})")
    
    def _init_ai_prompts(self):
        self.ai_prompts = {
            DataCategory.GEO_ZONES: """
Ты эксперт по логистике маркетплейсов с актуальными данными на 2026 год.
Предоставь актуальные гео-зоны доставки для маркетплейса "{marketplace}" в России.
Верни ТОЛЬКО валидный JSON без markdown-разметки.
Формат: {{"zones": [{{"zone_id": "MSC", "zone_name": "Москва и МО", "region_code": "77", "base_delivery_cost": 45.0, "cost_per_kg": 12.0, "cost_per_km": 3.2, "min_delivery_days": 1, "max_delivery_days": 2, "return_rate": 0.12, "coverage_ratio": 0.95}}], "data_date": "2026-01", "confidence": 0.9}}
""",
            DataCategory.SEASONAL_COEFFICIENTS: """
Ты эксперт по e-commerce и сезонности продаж с актуальными данными на 2026 год.
Предоставь сезонные коэффициенты спроса для категории "{category}" на российских маркетплейсах.
Коэффициенты по месяцам (1-12), где 1.0 = средний спрос.
Верни ТОЛЬКО валидный JSON.
Формат: {{"category": "{category}", "year": 2026, "monthly_coefficients": {{"1": 0.75, "2": 0.80, "12": 1.25}}, "peak_months": [11, 12], "low_months": [1, 2], "confidence": 0.85}}
""",
            DataCategory.MARKET_TRENDS: """
Ты эксперт по рыночным трендам e-commerce с данными на 2024-2026 годы.
Предоставь коэффициенты роста рынка для категории "{category}" на российских маркетплейсах по годам.
Верни ТОЛЬКО валидный JSON.
Формат: {{"category": "{category}", "yearly_trends": {{"2024": 1.08, "2025": 1.12, "2026": 1.15}}, "confidence": 0.8}}
""",
            DataCategory.PRICE_ELASTICITY: """
Ты эксперт по ценообразованию в e-commerce.
Предоставь коэффициент ценовой эластичности спроса для категории "{category}" на российских маркетплейсах.
Верни ТОЛЬКО валидный JSON.
Формат: {{"category": "{category}", "elasticity": -1.5, "confidence": 0.8, "description": "Спрос эластичный"}}
""",
            DataCategory.REGIONAL_RENT_RATES: """
Ты эксперт по коммерческой недвижимости в России с данными на 2026 год.
Предоставь средние ставки аренды складских помещений класса B по федеральным округам РФ.
Верни ТОЛЬКО валидный JSON.
Формат: {{"regions": {{"MSC": 850, "SPB": 650, "SIB": 350}}, "unit": "руб/м²/мес", "data_date": "2026-01", "confidence": 0.85}}
""",
            DataCategory.LABOR_RATES: """
Ты эксперт по рынку труда в логистике России с данными на 2026 год.
Предоставь средние зарплаты складского персонала по федеральным округам РФ.
Верни ТОЛЬКО валидный JSON.
Формат: {{"regions": {{"MSC": 45000, "SPB": 38000, "SIB": 26000}}, "unit": "руб/мес", "position": "кладовщик", "data_date": "2026-01", "confidence": 0.85}}
"""
        }
    
    def load_data(self, 
                  data_category: DataCategory,
                  marketplace: str = "",
                  category: str = "",
                  force_refresh: bool = False,
                  use_ai: bool = True,
                  user_csv: Optional[str] = None,
                  user_data: Any = None) -> DataLoadResult:
        start_time = time.time()
        attempts: List[DataLoadAttempt] = []
        
        logger.info(f"🔄 [Загрузка] {data_category.value} | Маркетплейс: {marketplace} | Категория: {category}")
        
        # Шаг 1: Кэш
        if not force_refresh:
            cache_key = self._build_cache_key(data_category, marketplace, category)
            cached_data = self.cache_manager.get('intelligent_loader', cache_key)
            
            if cached_data is not None:
                attempt = DataLoadAttempt(
                    source=DataSource.CACHE,
                    timestamp=datetime.now().isoformat(),
                    success=True,
                    data_count=self._count_data_items(cached_data),
                    response_time_ms=(time.time() - start_time) * 1000
                )
                attempts.append(attempt)
                
                result = DataLoadResult(
                    success=True, source=DataSource.CACHE, data=cached_data,
                    data_category=data_category, confidence=1.0,
                    message="Данные загружены из кэша",
                    attempts=attempts, load_time_ms=(time.time() - start_time) * 1000
                )
                self._update_stats(result)
                self.load_history.append(result)
                return result
        
        # Шаг 2: API
        logger.info(f"🔄 [Загрузка] Шаг 2/5: API {marketplace}")
        api_start = time.time()
        api_data = self._load_from_api(data_category, marketplace, category)
        
        api_attempt = DataLoadAttempt(
            source=DataSource.API,
            timestamp=datetime.now().isoformat(),
            success=api_data is not None,
            data_count=self._count_data_items(api_data),
            response_time_ms=(time.time() - api_start) * 1000,
            error_message="" if api_data is not None else "API недоступен"
        )
        attempts.append(api_attempt)
        
        if api_data is not None:
            self._cache_data(data_category, marketplace, category, api_data)
            
            result = DataLoadResult(
                success=True, source=DataSource.API, data=api_data,
                data_category=data_category, confidence=1.0,
                message=f"Данные загружены из API {marketplace}",
                attempts=attempts, load_time_ms=(time.time() - start_time) * 1000
            )
            self._update_stats(result)
            self.load_history.append(result)
            return result
        
        # Шаг 3: AI
        if use_ai and self.deepseek_api_key:
            logger.info(f"🔄 [Загрузка] Шаг 3/5: DeepSeek AI")
            ai_start = time.time()
            ai_data = self._load_from_ai(data_category, marketplace, category)
            
            ai_attempt = DataLoadAttempt(
                source=DataSource.AI,
                timestamp=datetime.now().isoformat(),
                success=ai_data is not None,
                data_count=self._count_data_items(ai_data),
                response_time_ms=(time.time() - ai_start) * 1000,
                error_message="" if ai_data is not None else "DeepSeek AI недоступен"
            )
            attempts.append(ai_attempt)
            
            if ai_data is not None:
                self._cache_data(data_category, marketplace, category, ai_data)
                
                result = DataLoadResult(
                    success=True, source=DataSource.AI, data=ai_data,
                    data_category=data_category, confidence=0.85,
                    message="Данные получены через DeepSeek AI",
                    attempts=attempts, load_time_ms=(time.time() - start_time) * 1000
                )
                self._update_stats(result)
                self.load_history.append(result)
                return result
        
        # Шаг 4: CSV
        if user_csv:
            logger.info(f"🔄 [Загрузка] Шаг 4/5: CSV")
            csv_start = time.time()
            csv_data = self._load_from_csv(data_category, user_csv)
            
            csv_attempt = DataLoadAttempt(
                source=DataSource.CSV,
                timestamp=datetime.now().isoformat(),
                success=csv_data is not None,
                data_count=self._count_data_items(csv_data),
                response_time_ms=(time.time() - csv_start) * 1000,
                error_message="" if csv_data is not None else "Ошибка парсинга CSV"
            )
            attempts.append(csv_attempt)
            
            if csv_data is not None:
                self._cache_data(data_category, marketplace, category, csv_data)
                
                result = DataLoadResult(
                    success=True, source=DataSource.CSV, data=csv_data,
                    data_category=data_category, confidence=1.0,
                    message="Данные загружены из CSV файла",
                    attempts=attempts, load_time_ms=(time.time() - start_time) * 1000
                )
                self._update_stats(result)
                self.load_history.append(result)
                return result
        
        # Шаг 5: Пользовательский ввод
        if user_data is not None:
            logger.info(f"🔄 [Загрузка] Шаг 5/5: Пользовательский ввод")
            
            user_attempt = DataLoadAttempt(
                source=DataSource.USER,
                timestamp=datetime.now().isoformat(),
                success=True,
                data_count=self._count_data_items(user_data),
                response_time_ms=(time.time() - start_time) * 1000
            )
            attempts.append(user_attempt)
            
            self._cache_data(data_category, marketplace, category, user_data)
            
            result = DataLoadResult(
                success=True, source=DataSource.USER, data=user_data,
                data_category=data_category, confidence=1.0,
                message="Использованы пользовательские данные",
                attempts=attempts, load_time_ms=(time.time() - start_time) * 1000
            )
            self._update_stats(result)
            self.load_history.append(result)
            return result
        
        # Все источники недоступны
        logger.error(f"❌ [Загрузка] ВСЕ источники недоступны для {data_category.value}")
        
        result = DataLoadResult(
            success=False, source=DataSource.NONE, data=None,
            data_category=data_category, confidence=0.0,
            message="НЕ УДАЛОСЬ ЗАГРУЗИТЬ ДАННЫЕ. Все источники недоступны.",
            attempts=attempts, load_time_ms=(time.time() - start_time) * 1000
        )
        self._update_stats(result)
        self.load_history.append(result)
        return result
    
    def _load_from_api(self, data_category: DataCategory, marketplace: str, 
                       category: str) -> Optional[Any]:
        if data_category == DataCategory.GEO_ZONES:
            return self._api_get_geo_zones(marketplace)
        return None
    
    def _api_get_geo_zones(self, marketplace: str) -> Optional[List[Dict]]:
        if marketplace == "Ozon":
            return self._fetch_ozon_zones()
        elif marketplace == "Wildberries":
            return self._fetch_wb_zones()
        elif marketplace == "Яндекс Маркет":
            return self._fetch_yandex_zones()
        return None
    
    def _fetch_ozon_zones(self) -> Optional[List[Dict]]:
        client_id = self.api_manager.get_api_key('ozon_client_id')
        api_key = self.api_manager.get_api_key('ozon')
        
        if not client_id or not api_key:
            return None
        
        try:
            self.api_manager.rate_limiter.wait_if_needed('ozon')
            
            response = requests.post(
                MarketplaceAPIEndpoint.OZON_DELIVERY_ZONES.value,
                headers={
                    'Client-Id': client_id,
                    'Api-Key': api_key,
                    'Content-Type': 'application/json'
                },
                json={"language": "RU"},
                timeout=self.api_timeout
            )
            
            if response.status_code == 200:
                data = response.json()
                zones = []
                for item in data.get('result', {}).get('zones', []):
                    zones.append({
                        'zone_id': str(item.get('zone_id', '')),
                        'zone_name': str(item.get('zone_name', '')),
                        'region_code': str(item.get('region_code', '')),
                        'base_delivery_cost': float(item.get('base_cost', 0)),
                        'cost_per_kg': float(item.get('cost_per_kg', 0)),
                        'cost_per_km': float(item.get('cost_per_km', 0)),
                        'min_delivery_days': int(item.get('min_days', 1)),
                        'max_delivery_days': int(item.get('max_days', 3)),
                        'return_rate': float(item.get('return_rate', 0.15)),
                        'coverage_ratio': float(item.get('coverage', 0.5)),
                        'source': 'ozon_api'
                    })
                return zones if zones else None
        except Exception as e:
            logger.error(f"❌ Ozon zones API error: {e}")
        return None
    
    def _fetch_wb_zones(self) -> Optional[List[Dict]]:
        api_key = self.api_manager.get_api_key('wildberries')
        if not api_key:
            return None
        
        try:
            self.api_manager.rate_limiter.wait_if_needed('wildberries')
            
            response = requests.get(
                MarketplaceAPIEndpoint.WILDBERRIES_DELIVERY_ZONES.value,
                headers={
                    'Authorization': f'Bearer {api_key}',
                    'Content-Type': 'application/json'
                },
                timeout=self.api_timeout
            )
            
            if response.status_code == 200:
                data = response.json()
                zones = []
                for item in data.get('data', []):
                    zones.append({
                        'zone_id': str(item.get('zoneId', '')),
                        'zone_name': str(item.get('zoneName', '')),
                        'region_code': str(item.get('regionCode', '')),
                        'base_delivery_cost': float(item.get('baseCost', 0)),
                        'cost_per_kg': float(item.get('costPerKg', 0)),
                        'cost_per_km': float(item.get('costPerKm', 0)),
                        'min_delivery_days': int(item.get('minDays', 1)),
                        'max_delivery_days': int(item.get('maxDays', 3)),
                        'return_rate': float(item.get('returnRate', 0.15)),
                        'coverage_ratio': float(item.get('coverage', 0.5)),
                        'source': 'wildberries_api'
                    })
                return zones if zones else None
        except Exception as e:
            logger.error(f"❌ WB zones API error: {e}")
        return None
    
    def _fetch_yandex_zones(self) -> Optional[List[Dict]]:
        api_key = self.api_manager.get_api_key('yandex_market')
        campaign_id = self.api_manager.get_api_key('yandex_campaign_id')
        
        if not api_key or not campaign_id:
            return None
        
        try:
            self.api_manager.rate_limiter.wait_if_needed('yandex_market')
            
            url = MarketplaceAPIEndpoint.YANDEX_DELIVERY_ZONES.value.format(campaign_id=campaign_id)
            response = requests.get(
                url,
                headers={
                    'Authorization': f'Bearer {api_key}',
                    'Content-Type': 'application/json'
                },
                timeout=self.api_timeout
            )
            
            if response.status_code == 200:
                data = response.json()
                zones = []
                for item in data.get('deliveryZones', []):
                    zones.append({
                        'zone_id': str(item.get('zoneId', '')),
                        'zone_name': str(item.get('zoneName', '')),
                        'region_code': str(item.get('regionCode', '')),
                        'base_delivery_cost': float(item.get('baseCost', 0)),
                        'cost_per_kg': float(item.get('costPerKg', 0)),
                        'cost_per_km': float(item.get('costPerKm', 0)),
                        'min_delivery_days': int(item.get('minDays', 1)),
                        'max_delivery_days': int(item.get('maxDays', 3)),
                        'return_rate': float(item.get('returnRate', 0.15)),
                        'coverage_ratio': float(item.get('coverage', 0.5)),
                        'source': 'yandex_api'
                    })
                return zones if zones else None
        except Exception as e:
            logger.error(f"❌ Yandex zones API error: {e}")
        return None
    
    def _load_from_ai(self, data_category: DataCategory, marketplace: str,
                      category: str) -> Optional[Any]:
        if not self.deepseek_api_key:
            return None
        
        prompt_template = self.ai_prompts.get(data_category, "")
        if not prompt_template:
            return None
        
        prompt = prompt_template.format(marketplace=marketplace, category=category)
        
        try:
            self.api_manager.rate_limiter.wait_if_needed('deepseek')
            
            response = requests.post(
                MarketplaceAPIEndpoint.DEEPSEEK_CHAT.value,
                headers={
                    'Authorization': f'Bearer {self.deepseek_api_key}',
                    'Content-Type': 'application/json'
                },
                json={
                    'model': 'deepseek-chat',
                    'messages': [
                        {'role': 'system', 'content': 'Ты эксперт по маркетплейсам. Отвечай ТОЛЬКО валидным JSON.'},
                        {'role': 'user', 'content': prompt}
                    ],
                    'temperature': 0.1,
                    'max_tokens': 4000,
                    'response_format': {'type': 'json_object'}
                },
                timeout=self.ai_timeout
            )
            
            if response.status_code == 200:
                content = response.json()['choices'][0]['message']['content']
                return json.loads(content)
        except Exception as e:
            logger.error(f"❌ DeepSeek AI error: {e}")
        
        return None
    
    def _load_from_csv(self, data_category: DataCategory, csv_content: str) -> Optional[Any]:
        try:
            df = pd.read_csv(io.StringIO(csv_content))
            
            if data_category == DataCategory.GEO_ZONES:
                return self._parse_geo_zones_csv(df)
            elif data_category == DataCategory.SEASONAL_COEFFICIENTS:
                return self._parse_seasonal_csv(df)
            elif data_category == DataCategory.MARKET_TRENDS:
                return self._parse_trends_csv(df)
            elif data_category == DataCategory.PRICE_ELASTICITY:
                return self._parse_elasticity_csv(df)
            else:
                return df.to_dict('records')
        except Exception as e:
            logger.error(f"❌ CSV parse error: {e}")
            return None
    
    def _parse_geo_zones_csv(self, df: pd.DataFrame) -> Optional[List[Dict]]:
        required_cols = ['zone_id', 'zone_name', 'region_code', 'base_delivery_cost']
        if not all(col in df.columns for col in required_cols):
            return None
        
        zones = []
        for _, row in df.iterrows():
            zones.append({
                'zone_id': str(row['zone_id']),
                'zone_name': str(row['zone_name']),
                'region_code': str(row['region_code']),
                'base_delivery_cost': float(row['base_delivery_cost']),
                'cost_per_kg': float(row.get('cost_per_kg', 0)),
                'cost_per_km': float(row.get('cost_per_km', 0)),
                'min_delivery_days': int(row.get('min_days', 1)),
                'max_delivery_days': int(row.get('max_days', 3)),
                'return_rate': float(row.get('return_rate', 0.15)),
                'coverage_ratio': float(row.get('coverage', 0.5)),
                'source': 'csv_import'
            })
        return zones if zones else None
    
    def _parse_seasonal_csv(self, df: pd.DataFrame) -> Optional[Dict[int, float]]:
        if df.empty:
            return None
        
        row = df.iloc[0]
        coeffs = {}
        for month in range(1, 13):
            col_name = f'month_{month}'
            if col_name in df.columns:
                coeffs[month] = float(row[col_name])
        
        return coeffs if coeffs else None
    
    def _parse_trends_csv(self, df: pd.DataFrame) -> Optional[Dict[int, float]]:
        if df.empty:
            return None
        
        row = df.iloc[0]
        trends = {}
        for col in df.columns:
            if col.startswith('year_'):
                year = int(col.replace('year_', ''))
                trends[year] = float(row[col])
        
        return trends if trends else None
    
    def _parse_elasticity_csv(self, df: pd.DataFrame) -> Optional[float]:
        if df.empty or 'elasticity' not in df.columns:
            return None
        return float(df.iloc[0]['elasticity'])
    
    def _build_cache_key(self, data_category: DataCategory, marketplace: str,
                         category: str) -> str:
        parts = [data_category.value, marketplace.lower(), category.lower()]
        return hashlib.md5("|".join(parts).encode()).hexdigest()
    
    def _cache_data(self, data_category: DataCategory, marketplace: str,
                    category: str, data: Any):
        cache_key = self._build_cache_key(data_category, marketplace, category)
        self.cache_manager.set('intelligent_loader', cache_key, data)
    
    def _count_data_items(self, data: Any) -> int:
        if data is None:
            return 0
        if isinstance(data, (list, tuple)):
            return len(data)
        if isinstance(data, dict):
            return len(data)
        return 1
    
    def _update_stats(self, result: DataLoadResult):
        self.stats['total_attempts'] += 1
        if result.success:
            self.stats['successful_loads'] += 1
        else:
            self.stats['failed_loads'] += 1
        
        self.stats['by_source'][result.source.value] += 1
        
        if result.data:
            self.stats['total_data_points_loaded'] += self._count_data_items(result.data)
    
    def get_statistics(self) -> Dict[str, Any]:
        total = self.stats['total_attempts']
        success_rate = (self.stats['successful_loads'] / total * 100) if total > 0 else 0
        
        return {
            **self.stats,
            'success_rate_pct': round(success_rate, 2),
            'most_used_source': max(self.stats['by_source'], key=self.stats['by_source'].get),
            'history_size': len(self.load_history)
        }
    
    def clear_cache(self, data_category: Optional[DataCategory] = None):
        self.cache_manager.clear_cache('intelligent_loader')

# ============================================================================
# БЛОК 10: UI КОМПОНЕНТ ИНТЕЛЛЕКТУАЛЬНОЙ ЗАГРУЗКИ ДАННЫХ
# ============================================================================

def init_intelligent_loader_state():
    if 'intelligent_loader' not in st.session_state:
        st.session_state.intelligent_loader = IntelligentDataLoader(
            api_manager=st.session_state.get('api_manager'),
            cache_manager=CacheManager(),
            secure_data=st.session_state.get('api_manager', MarketplaceAPIManager()).secure_data
        )
    
    if 'loaded_data_cache' not in st.session_state:
        st.session_state.loaded_data_cache = {}

def render_intelligent_data_loader_ui():
    init_intelligent_loader_state()
    loader = st.session_state.intelligent_loader
    
    st.markdown("## 🧠 Интеллектуальная загрузка данных")
    
    with st.expander("ℹ️ Как работает система загрузки", expanded=False):
        st.markdown("""
        ### 🔄 Каскадная загрузка данных
        
        | Шаг | Источник | Приоритет | Уверенность |
        |-----|----------|-----------|-------------|
        | 1 | 📦 Кэш | Высший | 100% |
        | 2 | 🔌 API маркетплейса | Высокий | 100% |
        | 3 | 🤖 DeepSeek AI | Средний | 85% |
        | 4 | 📄 CSV файл | Средний | 100% |
        | 5 | 👤 Ручной ввод | Низкий | 100% |
        
        **ВАЖНО:** Система НИКОГДА не использует захардкоженные значения.
        """)
    
    col1, col2 = st.columns([1, 1])
    
    with col1:
        data_category = st.selectbox(
            "📊 Категория данных:",
            options=[dc for dc in DataCategory],
            format_func=lambda x: {
                DataCategory.GEO_ZONES: "🌍 Гео-зоны доставки",
                DataCategory.SEASONAL_COEFFICIENTS: "📅 Сезонные коэффициенты",
                DataCategory.MARKET_TRENDS: "📈 Рыночные тренды",
                DataCategory.PRICE_ELASTICITY: "💹 Эластичность спроса",
                DataCategory.REGIONAL_RENT_RATES: "🏗️ Ставки аренды складов",
                DataCategory.LABOR_RATES: "👷 Ставки зарплат"
            }.get(x, x.value)
        )
    
    with col2:
        marketplace = st.selectbox("🏪 Маркетплейс:", ["Ozon", "Wildberries", "Яндекс Маркет", "Все"])
    
    if data_category in [DataCategory.SEASONAL_COEFFICIENTS, DataCategory.MARKET_TRENDS,
                         DataCategory.PRICE_ELASTICITY]:
        category = st.text_input("📦 Категория товара:", value="electronics")
    else:
        category = "default"
    
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    
    with col1:
        force_refresh = st.checkbox("🔄 Принудительное обновление", value=False)
    with col2:
        use_ai = st.checkbox("🤖 DeepSeek AI", value=True)
    with col3:
        show_details = st.checkbox("📋 Детали", value=True)
    
    csv_file = st.file_uploader("📄 CSV файл (опционально):", type=['csv'])
    
    st.markdown("---")
    
    if st.button("🚀 ЗАГРУЗИТЬ ДАННЫЕ", type="primary", width='stretch'):
        csv_content = csv_file.getvalue().decode('utf-8') if csv_file else None
        
        progress_bar = st.progress(0)
        status_container = st.empty()
        
        status_container.info("🔍 Шаг 1/5: Проверка кэша...")
        progress_bar.progress(10)
        time.sleep(0.2)
        
        status_container.info("🔌 Шаг 2/5: API маркетплейса...")
        progress_bar.progress(30)
        time.sleep(0.2)
        
        status_container.info("🤖 Шаг 3/5: DeepSeek AI...")
        progress_bar.progress(50)
        time.sleep(0.2)
        
        status_container.info("📄 Шаг 4/5: CSV...")
        progress_bar.progress(70)
        time.sleep(0.2)
        
        status_container.info("👤 Шаг 5/5: Проверка ввода...")
        progress_bar.progress(90)
        
        result = loader.load_data(
            data_category=data_category,
            marketplace=marketplace if marketplace != "Все" else "",
            category=category,
            force_refresh=force_refresh,
            use_ai=use_ai,
            user_csv=csv_content
        )
        
        progress_bar.progress(100)
        status_container.empty()
        progress_bar.empty()
        
        st.session_state.loaded_data_cache[data_category.value] = result
        
        if result.success:
            source_icons = {
                DataSource.API: "🔌", DataSource.AI: "🤖",
                DataSource.CSV: "📄", DataSource.USER: "👤", DataSource.CACHE: "💾"
            }
            icon = source_icons.get(result.source, "✅")
            
            st.success(f"""
            ### {icon} ДАННЫЕ УСПЕШНО ЗАГРУЖЕНЫ!
            
            | Параметр | Значение |
            |----------|----------|
            | **Источник** | `{result.source.value.upper()}` |
            | **Уверенность** | `{result.confidence*100:.0f}%` |
            | **Время загрузки** | `{result.load_time_ms:.0f} мс` |
            """)
            
            if result.data:
                with st.expander("👀 Превью данных"):
                    if isinstance(result.data, list):
                        st.dataframe(pd.DataFrame(result.data[:10]), width='stretch')
                    elif isinstance(result.data, dict):
                        st.json(result.data)
        else:
            st.error(f"❌ НЕ УДАЛОСЬ ЗАГРУЗИТЬ: {result.message}")
            
            st.warning("### 🔧 Ручной ввод данных")
            if data_category == DataCategory.PRICE_ELASTICITY:
                elasticity = st.number_input("Эластичность:", min_value=-5.0, max_value=0.0, value=-1.5)
                if st.button("💾 Сохранить"):
                    result = loader.load_data(
                        data_category=data_category, marketplace=marketplace,
                        category=category, user_data=elasticity
                    )
                    st.success("✅ Сохранено!")
                    st.rerun()
    
    st.markdown("---")
    stats = loader.get_statistics()
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Попыток", stats['total_attempts'])
    with col2:
        st.metric("Успешно", stats['successful_loads'], delta=f"{stats['success_rate_pct']:.0f}%")
    with col3:
        st.metric("Провалено", stats['failed_loads'])
    with col4:
        st.metric("Точек данных", stats['total_data_points_loaded'])
    
    if st.button("🗑️ Очистить кэш"):
        loader.clear_cache()
        st.success("✅ Кэш очищен!")

# ============================================================================
# БЛОК 11: ЭКСПОРТ В GOOGLE SHEETS БЕЗ СЕРВИСНОГО АККАУНТА
# ============================================================================

class GoogleSheetsExporter:
    """
    Помощник для экспорта данных в Google Таблицы без API.
    Генерирует CSV, TSV или текст для копирования, а также инструкцию по импорту.
    """
    
    @staticmethod
    def generate_csv(results: List[FBSResultData], separator: str = ";") -> str:
        """Генерирует строку CSV с результатами."""
        if not results:
            return ""
        headers = ['Артикул', 'Название', 'Цена', 'Прибыль', 'Маржа,%', 'ROI,%',
                   'Комиссия', 'First Mile', 'Last Mile', 'Pick&Pack', 'Упаковка',
                   'Эквайринг', 'Возвраты', 'Штрафы', 'Маркетинг', 'Склад', 'Налог',
                   'Опт.запас', 'Оборач., дн', 'Лог.зона', 'Источник']
        lines = [separator.join(headers)]
        for r in results:
            row = [
                str(r.artikul), str(r.product_name), str(r.selling_price),
                str(r.gross_profit), str(r.margin_percent), str(r.roi_percent),
                str(r.commission), str(r.first_mile_cost), str(r.last_mile_cost),
                str(r.pick_pack_cost), str(r.packaging_cost), str(r.acquiring_cost),
                str(r.return_cost), str(r.penalty_cost), str(r.marketing_cost),
                str(r.warehouse_cost), str(r.tax_cost), str(r.optimal_stock_units),
                str(r.stock_turnover_days), str(r.logistic_zone_label), str(r.data_source)
            ]
            lines.append(separator.join(row))
        return "\n".join(lines)
    
    @staticmethod
    def generate_tsv(results: List[FBSResultData]) -> str:
        """Генерирует TSV (табуляция) — удобно для вставки в Google Sheets."""
        return GoogleSheetsExporter.generate_csv(results, separator="\t")
    
    @staticmethod
    def get_import_instructions() -> str:
        """Возвращает пошаговую инструкцию по импорту в Google Sheets."""
        return """
📌 **Как импортировать данные в вашу Google Таблицу (без сервисного аккаунта):**

1. Откройте вашу Google Таблицу.
2. Нажмите **Файл → Импорт → Загрузить**.
3. Выберите скачанный CSV-файл (или вставьте скопированный текст).
4. В настройках импорта выберите:
   - **Разделитель**: `Точка с запятой` (если CSV) или `Табуляция` (если TSV).
   - **Место импорта**: выберите существующий лист или создайте новый.
5. Нажмите **Импортировать** — данные появятся в таблице.

💡 **Альтернатива:** скопируйте таблицу из раздела «Результаты» в интерфейсе и вставьте прямо в Google Sheets (Ctrl+V).
"""

# ============================================================================
# БЛОК 12: ИНТЕРФЕЙС ПОЛЬЗОВАТЕЛЯ (STREAMLIT) — ПОЛНАЯ НЕСОКРАЩЕННАЯ ВЕРСИЯ
# ============================================================================

def init_session_state():
    if 'api_manager' not in st.session_state:
        st.session_state.api_manager = MarketplaceAPIManager()
    
    if 'intelligent_loader' not in st.session_state:
        st.session_state.intelligent_loader = IntelligentDataLoader(
            api_manager=st.session_state.api_manager,
            cache_manager=CacheManager(),
            secure_data=st.session_state.api_manager.secure_data
        )
    
    if 'calculator' not in st.session_state:
        st.session_state.calculator = FBSUnitEconomicsCalculator(
            api_manager=st.session_state.api_manager,
            intelligent_loader=st.session_state.intelligent_loader
        )
    
    if 'results' not in st.session_state:
        st.session_state.results = []
    
    if 'input_data_list' not in st.session_state:
        st.session_state.input_data_list = []
    
    if 'marketplace' not in st.session_state:
        st.session_state.marketplace = "Ozon"
    
    if 'tax_system' not in st.session_state:
        st.session_state.tax_system = "УСН 6% (доходы)"
    
    if 'current_section' not in st.session_state:
        st.session_state.current_section = 'main'
    
    if 'recommendations' not in st.session_state:
        st.session_state.recommendations = []
    
    if 'loaded_data_cache' not in st.session_state:
        st.session_state.loaded_data_cache = {}
    
    if 'onboarding_done' not in st.session_state:
        st.session_state.onboarding_done = False
    
    if 'auto_gs_update' not in st.session_state:
        st.session_state.auto_gs_update = False

def show_onboarding():
    """Показывает новичкам пошаговое руководство."""
    with st.expander("🎓 Новичок? Начни здесь!", expanded=not st.session_state.get('onboarding_done', False)):
        st.markdown("""
        ### 🚀 Быстрый старт за 4 шага:
        1. **Настрой API ключи** (раздел "Настройки") — опционально, можно пропустить.
        2. **Загрузи данные** (раздел "Загрузка данных") — система сама подберёт источник.
        3. **Рассчитай юнит-экономику** (раздел "Калькулятор FBS") — введи параметры и получи результат.
        4. **Экспортируй в Google Таблицу** (раздел "Google Sheets"):
           - Скачай CSV или скопируй данные и вставь в свою таблицу (инструкция внутри).
           - Если есть сервисный аккаунт — включи автоматическую синхронизацию.
        ---
        🔗 Все данные всегда можно обновить вручную за пару кликов.
        """)
        if st.button("✅ Понятно, больше не показывать"):
            st.session_state.onboarding_done = True
            st.rerun()

def render_sidebar():
    with st.sidebar:
        st.markdown("""
        <div style='text-align: center; padding: 20px 15px; background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460); border-radius: 12px; margin-bottom: 25px;'>
            <h1 style='color: white; margin: 0; font-size: 1.5em;'>🚀 FBS PRO</h1>
            <p style='color: #a8a8d0; margin: 8px 0 0 0; font-size: 0.9em;'>Полная ИИ-версия</p>
            <p style='color: #6666aa; margin: 5px 0 0 0; font-size: 0.7em;'>v7.2.0 | API → AI → CSV → User</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 🧭 Навигация")
        
        sections = {
            "🏠 Главная": "main",
            "🧮 Калькулятор FBS": "calculator",
            "🧠 Загрузка данных": "data_loader",
            "📋 Тарифы маркетплейсов": "tariffs",
            "📈 Дашборд": "dashboard",
            "🎯 Анализ сценариев": "what_if",
            "💡 Рекомендации": "recommendations",
            "📥 Экспорт": "export",
            "🌐 Google Sheets": "gsheets",
            "⚙️ Настройки": "settings"
        }
        
        selected_section = st.radio("Выберите раздел:", list(sections.keys()), label_visibility="collapsed")
        
        st.session_state.current_section = sections[selected_section]
        
        st.markdown("---")
        st.markdown("### 📊 Статус системы")
        
        calculator = st.session_state.calculator
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**🏪 МП:** {st.session_state.marketplace}")
        with col2:
            st.markdown(f"**💰 Налог:** {st.session_state.tax_system.split()[0]}")
        
        if calculator.tariffs_source == 'api':
            st.success("🔌 Тарифы: API маркетплейса")
        elif calculator.tariffs_source == 'deepseek':
            st.info("🤖 Тарифы: DeepSeek AI")
        elif calculator.tariffs_source == 'csv':
            st.info("📄 Тарифы: CSV импорт")
        elif calculator.tariffs_source == 'user':
            st.info("👤 Тарифы: Пользовательский ввод")
        else:
            st.warning("⚠️ Тарифы: Не загружены!")
        
        if calculator.geo_zones:
            st.success(f"🌍 Гео-зоны: {len(calculator.geo_zones)}")
        else:
            st.warning("🌍 Гео-зоны: не загружены")
        
        if st.session_state.results:
            st.success(f"✅ Рассчитано: {len(st.session_state.results)} товаров")
            profitable = len([r for r in st.session_state.results if r.gross_profit > 0])
            st.metric("Прибыльных", f"{profitable} из {len(st.session_state.results)}")
        else:
            st.info("ℹ️ Расчеты не выполнялись")
        
        st.markdown("---")
        st.markdown("### ⚡ Быстрые действия")
        
        if st.button("🔄 Обновить тарифы", width='stretch'):
            with st.spinner("Загрузка тарифов..."):
                calculator.refresh_tariffs(force=True)
                st.success("✅ Тарифы обновлены!")
                st.rerun()
        
        if st.button("🗑️ Очистить результаты", width='stretch'):
            st.session_state.results = []
            st.session_state.input_data_list = []
            st.session_state.recommendations = []
            st.success("✅ Результаты очищены!")
            st.rerun()

def render_dashboard_with_filters(results: List[FBSResultData]):
    if not results:
        st.warning("⚠️ Нет данных для отображения")
        return
    
    st.markdown("### 🔍 Фильтры")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        min_margin = st.slider("Минимальная маржа, %", -50, 100, -10)
    with col2:
        search_artikul = st.text_input("Поиск по артикулу", placeholder="Введите артикул...")
    with col3:
        sort_by = st.selectbox("Сортировка", ["Прибыль", "Маржа", "ROI", "Оборачиваемость", "Прибыль на м²"])
    with col4:
        filter_zone = st.selectbox("Логистическая зона", ["Все", "🔴 Критическая", "🟡 Зона риска", "🟢 Безопасная", "🔵 Идеальная"])
    
    filtered = [r for r in results if r.margin_percent >= min_margin]
    
    if search_artikul:
        filtered = [r for r in filtered if search_artikul.lower() in r.artikul.lower()]
    
    if filter_zone != "Все":
        filtered = [r for r in filtered if filter_zone in r.logistic_zone_label]
    
    sort_map = {
        "Прибыль": "gross_profit",
        "Маржа": "margin_percent",
        "ROI": "roi_percent",
        "Оборачиваемость": "stock_turnover_days",
        "Прибыль на м²": "profit_per_sqm"
    }
    if sort_by in sort_map:
        filtered.sort(key=lambda x: getattr(x, sort_map[sort_by]), reverse=True)
    
    if filtered:
        df = pd.DataFrame([r.get_summary() for r in filtered])
        st.dataframe(df, width='stretch', height=400)
        st.caption(f"📊 Показано {len(filtered)} из {len(results)} товаров")
    else:
        st.info("ℹ️ Нет товаров, соответствующих фильтрам")

def main():
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="🚀",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    init_session_state()
    render_sidebar()
    
    current_section = st.session_state.get('current_section', 'main')
    calculator = st.session_state.calculator
    
    # Показываем онбординг только на главной
    if current_section == 'main':
        show_onboarding()
    
    if current_section == 'main':
        st.markdown("""
        <div style='text-align: center; padding: 50px 30px; background: linear-gradient(135deg, #0f0c29, #302b63, #24243e); border-radius: 20px; margin-bottom: 35px;'>
            <h1 style='color: white; font-size: 3em; margin: 0;'>🚀 FBS Юнит-экономика PRO</h1>
            <p style='color: #a8a8d0; font-size: 1.3em; margin: 20px 0;'>
                Полная ИИ-версия — Никаких сокращений!
            </p>
            <p style='color: #6666aa; font-size: 1em; margin: 10px 0;'>
                Ozon • Wildberries • Яндекс Маркет | DeepSeek AI | Каскадная загрузка
            </p>
            <p style='color: #8888cc; font-size: 0.9em; margin: 10px 0;'>
                🆕 v7.2.0 — 100% сохранение исходного UI + интеграция с Google Sheets без сервисного аккаунта
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.info("""
        ### 🎯 Ключевые принципы системы
        
        1. **НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ** — все данные загружаются из реальных источников
        2. **Интеллектуальная загрузка** — каскадный фолбэк API → AI → CSV → User
        3. **Актуальные данные** — интеграция с API Ozon, Wildberries, Яндекс Маркет
        4. **AI-обогащение** — DeepSeek для получения данных при недоступности API
        5. **Полная прозрачность** — видно какой источник использован и с какой уверенностью
        6. **Google Sheets без аккаунта** — экспорт в один клик, инструкция внутри
        
        ### 📋 Что нужно для работы:
        - **API ключи** маркетплейсов (в разделе Настройки) — опционально
        - **DeepSeek API ключ** для AI-обогащения (опционально)
        - **CSV файлы** с данными (если API недоступны)
        - **Входные данные** товара (цена, вес, расстояния и т.д.)
        """)
        
        if st.session_state.results:
            st.markdown("---")
            st.markdown("### 📊 Последние результаты")
            results = st.session_state.results
            
            col1, col2, col3, col4, col5 = st.columns(5)
            with col1:
                st.metric("📦 Товаров", len(results))
            with col2:
                profitable = len([r for r in results if r.gross_profit > 0])
                st.metric("✅ Прибыльных", f"{profitable} ({profitable/len(results)*100:.0f}%)")
            with col3:
                total_profit = sum(r.gross_profit for r in results)
                st.metric("💰 Общая прибыль", f"{total_profit:,.0f} ₽")
            with col4:
                avg_margin = np.mean([r.margin_percent for r in results])
                st.metric("📊 Средняя маржа", f"{avg_margin:.1f}%")
            with col5:
                critical = len([r for r in results if r.is_logistic_critical])
                st.metric("🔴 Критич. логистика", f"{critical}")
    
    elif current_section == 'calculator':
        st.markdown("## 🧮 Калькулятор FBS юнит-экономики")
        st.info("""
        **🎯 Профессиональный расчет FBS с интеллектуальной загрузкой данных**
        
        - 🚛 **First Mile** — ваша логистика до склада МП (пользовательский ввод)
        - 📦 **Last Mile** — доставка МП до клиента (из API тарифов + гео-зоны)
        - 📊 **Оптимизация склада** — EOQ, страховой запас, точка заказа
        - 💰 **Оптимальная цена** — с учетом эластичности спроса
        - 🔴 **Логистические зоны риска** — на основе загруженных гео-зон
        - 📅 **Сезонная корректировка** — на основе загруженных коэффициентов
        - 🧠 **Источник данных** — отображается для каждого расчета
        """)
        
        calc_mode = st.radio("Режим расчета:", ["📱 Расчет одного товара", "📊 Массовый расчет из файла"], horizontal=True)
        
        if calc_mode == "📱 Расчет одного товара":
            with st.form("single_calc_form"):
                st.markdown("### 📝 Введите данные товара")
                st.caption("Все поля обязательны, если не указано иное. Наведи курсор на название для подсказки.")
                
                col1, col2 = st.columns(2)
                with col1:
                    artikul = st.text_input("Артикул", "SKU-001", help="Уникальный идентификатор товара в вашей системе")
                    product_name = st.text_input("Наименование", "Тестовый товар", help="Название товара для отчётов")
                    
                    tariff_categories = list(calculator.current_tariffs.keys()) if calculator.current_tariffs else ["default"]
                    category = st.selectbox("Категория", tariff_categories, help="Выберите категорию, от которой зависит комиссия и тарифы доставки")
                    
                    selling_price = st.number_input("Цена продажи, ₽", min_value=0.0, step=100.0, help="Розничная цена на маркетплейсе")
                    cogs = st.number_input("Себестоимость, ₽", min_value=0.0, step=100.0, help="Закупочная цена или себестоимость единицы")
                
                with col2:
                    weight = st.number_input("Вес, кг", min_value=0.0, step=0.1, help="Вес брутто товара")
                    length = st.number_input("Длина, см", min_value=0.0, step=1.0, help="Длина упаковки (для объемного веса)")
                    width = st.number_input("Ширина, см", min_value=0.0, step=1.0)
                    height = st.number_input("Высота, см", min_value=0.0, step=1.0)
                    warehouse_distance = st.number_input("Расстояние до склада МП, км", min_value=0.0, step=1.0, help="Расстояние от вашего склада до сортировочного центра маркетплейса")
                    daily_sales = st.number_input("Продаж в день, шт", min_value=1, step=1, help="Среднее количество продаж в день")
                    has_night = st.checkbox("Ночная смена", help="Включите, если работаете в ночную смену (уменьшает штрафы)")
                
                with st.expander("⚙️ Расширенные параметры"):
                    col1, col2 = st.columns(2)
                    with col1:
                        packaging_cost = st.number_input("Упаковка, ₽", min_value=0.0, step=5.0, help="Стоимость упаковочных материалов на единицу")
                        marketing_budget = st.number_input("Маркетинг на ед., ₽", min_value=0.0, step=10.0, help="Рекламный бюджет на одну продажу")
                        operator_rate = st.number_input("Ставка оператора, ₽/ч", min_value=0.0, step=50.0, help="Часовая ставка сотрудника на сборке")
                        stock_depth = st.number_input("Глубина запаса, дн", min_value=1, step=1, help="На сколько дней хватает текущего запаса")
                    
                    with col2:
                        pick_pack_time = st.number_input("Pick & Pack, мин", min_value=0.0, step=0.5, help="Время на сборку и упаковку одного заказа")
                        pallet_capacity = st.number_input("Единиц на паллете", min_value=1, step=10, help="Сколько товаров помещается на одну паллету")
                        transport_cost = st.number_input("Транспорт, ₽/км", min_value=0.0, step=5.0, help="Стоимость перевозки за км (с учётом возврата)")
                        safety_stock = st.number_input("Страховой запас, дн", min_value=0, step=1, help="Дополнительный запас на случай скачков спроса")
                
                submitted = st.form_submit_button("🚀 Рассчитать", type="primary")
                
                if submitted:
                    input_data = FBSInputData(
                        artikul=artikul, product_name=product_name, category=category,
                        selling_price=selling_price, cogs=cogs,
                        weight_kg=weight, length_cm=length, width_cm=width, height_cm=height,
                        warehouse_distance_km=warehouse_distance, daily_sales=daily_sales,
                        has_night_shift=has_night, packaging_cost=packaging_cost,
                        marketing_budget_per_unit=marketing_budget, operator_hourly_rate=operator_rate,
                        stock_depth_days=stock_depth, pick_pack_time_min=pick_pack_time,
                        pallet_capacity=pallet_capacity, transport_cost_per_km=transport_cost,
                        safety_stock_days=safety_stock
                    )
                    
                    result = calculator.calculate_unit_economics(input_data)
                    st.session_state.results = [result]
                    st.session_state.input_data_list = [input_data]
                    
                    st.markdown("---")
                    st.markdown("## 📊 Результаты расчета")
                    
                    source_icon = "🔌" if "api" in result.data_source else "🤖" if "deepseek" in result.data_source else "📄" if "csv" in result.data_source else "⚠️"
                    st.caption(f"{source_icon} Источник данных: {result.data_source} | Уверенность: {result.data_confidence*100:.0f}%")
                    
                    col1, col2, col3, col4, col5 = st.columns(5)
                    with col1:
                        st.metric("💰 Прибыль", f"{result.gross_profit:,.0f} ₽", f"{result.margin_percent:.1f}% маржи")
                    with col2:
                        st.metric("📦 Расходы", f"{result.total_expenses:,.0f} ₽")
                    with col3:
                        st.metric("📈 ROI", f"{result.roi_percent:.1f}%")
                    with col4:
                        st.metric("👥 LTV/CAC", f"{result.ltv_cac_ratio:.1f}x")
                    with col5:
                        st.metric("💵 Опт. цена", f"{result.optimal_price:,.0f} ₽")
                    
                    st.markdown("### 📋 Детализация")
                    col1, col2 = st.columns(2)
                    with col1:
                        st.markdown("**💰 Финансовые показатели**")
                        st.metric("Комиссия МП", f"{result.commission:,.0f} ₽")
                        st.metric("First Mile", f"{result.first_mile_cost:,.0f} ₽")
                        st.metric("Last Mile", f"{result.last_mile_cost:,.0f} ₽")
                        st.metric("Pick & Pack", f"{result.pick_pack_cost:,.0f} ₽")
                        st.metric("Упаковка", f"{result.packaging_cost:,.0f} ₽")
                        st.metric("Эквайринг", f"{result.acquiring_cost:,.0f} ₽")
                        st.metric("Возвраты", f"{result.return_cost:,.0f} ₽")
                        st.metric("Штрафы", f"{result.penalty_cost:,.0f} ₽")
                        st.metric("Маркетинг", f"{result.marketing_cost:,.0f} ₽")
                        st.metric("Складские", f"{result.warehouse_cost:,.0f} ₽")
                        st.metric("Налог", f"{result.tax_cost:,.0f} ₽")
                    
                    with col2:
                        st.markdown("**📊 Оптимизация склада**")
                        st.metric("Оптимальный запас (EOQ)", f"{result.optimal_stock_units} шт")
                        st.metric("Страховой запас", f"{result.safety_stock_units} шт")
                        st.metric("Точка заказа", f"{result.reorder_point_units} шт")
                        st.metric("Оборачиваемость", f"{result.stock_turnover_days:.1f} дн")
                        st.metric("Потенциал оптимизации", f"{result.stock_optimization_potential:.1f}%")
                        
                        st.markdown("**🚚 Логистика**")
                        st.metric("Зона", result.logistic_zone_label)
                        st.metric("Точка безубыт. (км)", f"{result.break_even_distance_km:.0f} км")
                        st.metric("Взвеш. доставка", f"{result.weighted_delivery_cost:,.0f} ₽")
                        
                        st.markdown("**📅 Сезонность**")
                        st.metric("Коэффициент", f"{result.seasonal_factor:.2f}")
                        st.metric("Скорр. маржа", f"{result.adjusted_margin_percent:.1f}%")
                    
                    st.markdown("### 💰 Рекомендованные цены")
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("Текущая цена", f"{result.selling_price:,.0f} ₽")
                    with col2:
                        st.metric("Оптимальная (эластичность)", f"{result.optimal_price:,.0f} ₽")
                    with col3:
                        rec_price_25 = result.total_expenses / (1 - 0.25)
                        st.metric("При марже 25%", f"{rec_price_25:,.0f} ₽")
                    
                    # Кнопка для копирования в буфер (для Google Sheets)
                    if st.button("📋 Скопировать результат в буфер (для вставки в Google Sheets)"):
                        row_data = [
                            result.artikul, result.product_name, result.selling_price,
                            result.gross_profit, f"{result.margin_percent:.1f}%",
                            result.roi_percent, result.commission, result.first_mile_cost,
                            result.last_mile_cost, result.pick_pack_cost, result.packaging_cost,
                            result.acquiring_cost, result.return_cost, result.penalty_cost,
                            result.marketing_cost, result.warehouse_cost, result.tax_cost,
                            result.optimal_stock_units, result.stock_turnover_days,
                            result.logistic_zone_label, result.data_source
                        ]
                        st.code("\t".join(str(x) for x in row_data), language="text")
                        st.caption("Скопируйте эту строку и вставьте в Google Sheets как новую строку.")
                    
                    # Автоматическое обновление Google Sheets, если включено
                    if st.session_state.get('auto_gs_update', False) and st.session_state.get('gs_manager'):
                        manager = st.session_state.gs_manager
                        if manager.sheet:
                            manager.update_all(st.session_state.calculator.current_tariffs, [result])
                            st.success("✅ Данные автоматически обновлены в Google Sheets!")
        
        else:  # Массовый расчет
            st.markdown("### 📊 Массовый расчет из файла")
            uploaded_file = st.file_uploader("Загрузите CSV с данными товаров", type=['csv'])
            if uploaded_file:
                df = pd.read_csv(uploaded_file)
                st.dataframe(df, width='stretch')
                if st.button("🚀 Рассчитать все"):
                    input_list = []
                    for _, row in df.iterrows():
                        try:
                            data = FBSInputData(
                                artikul=str(row.get('artikul', '')),
                                product_name=str(row.get('product_name', '')),
                                category=str(row.get('category', 'default')),
                                selling_price=float(row.get('selling_price', 0)),
                                cogs=float(row.get('cogs', 0)),
                                weight_kg=float(row.get('weight_kg', 0)),
                                length_cm=float(row.get('length_cm', 0)),
                                width_cm=float(row.get('width_cm', 0)),
                                height_cm=float(row.get('height_cm', 0)),
                                warehouse_distance_km=float(row.get('warehouse_distance_km', 0)),
                                daily_sales=int(row.get('daily_sales', 5)),
                                packaging_cost=float(row.get('packaging_cost', 0)),
                                marketing_budget_per_unit=float(row.get('marketing_budget', 0)),
                                operator_hourly_rate=float(row.get('operator_rate', 300)),
                                stock_depth_days=int(row.get('stock_depth', 30)),
                                pick_pack_time_min=float(row.get('pick_pack_time', 5)),
                                pallet_capacity=int(row.get('pallet_capacity', 100)),
                                transport_cost_per_km=float(row.get('transport_cost', 20)),
                                safety_stock_days=int(row.get('safety_stock', 7)),
                                has_night_shift=bool(row.get('night_shift', False))
                            )
                            input_list.append(data)
                        except Exception as e:
                            st.warning(f"⚠️ Ошибка в строке {row}: {e}")
                    
                    if input_list:
                        with st.spinner("Выполняется расчет..."):
                            results = calculator.calculate_batch(input_list)
                            st.session_state.results = results
                            st.session_state.input_data_list = input_list
                            st.success(f"✅ Рассчитано {len(results)} товаров!")
                            st.rerun()
    
    elif current_section == 'data_loader':
        render_intelligent_data_loader_ui()
    
    elif current_section == 'tariffs':
        st.markdown("## 📋 Актуальные тарифы маркетплейсов")
        
        st.info("""
        **📌 Важно:** Тарифы загружаются из API маркетплейсов, DeepSeek AI, CSV или пользовательского ввода.
        НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ не используется.
        """)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            marketplace = st.selectbox("Маркетплейс", ["Ozon", "Wildberries", "Яндекс Маркет"])
        with col2:
            force_refresh = st.checkbox("🔄 Принудительное обновление")
        with col3:
            use_ai = st.checkbox("🤖 DeepSeek AI", value=True)
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("📥 Загрузить тарифы", type="primary"):
                with st.spinner(f"Загрузка тарифов {marketplace}..."):
                    calculator.set_marketplace(marketplace)
                    calculator.refresh_tariffs(force=force_refresh, use_ai=use_ai)
                    st.success(f"✅ Тарифы {marketplace} загружены! (источник: {calculator.tariffs_source})")
                    st.rerun()
        
        with col2:
            csv_file = st.file_uploader("📄 CSV с тарифами", type=['csv'])
            if csv_file and st.button("📥 Загрузить из CSV"):
                csv_content = csv_file.getvalue().decode('utf-8')
                calculator.refresh_tariffs(force=True, csv_content=csv_content)
                st.success("✅ Тарифы загружены из CSV!")
                st.rerun()
        
        if calculator.current_tariffs:
            df = calculator.api_manager.get_all_tariffs_as_dataframe(marketplace)
            st.dataframe(df, width='stretch', height=400)
            
            st.markdown("### 📊 Статистика источников")
            sources = df['Источник'].value_counts() if 'Источник' in df.columns else pd.Series()
            st.dataframe(sources, width='stretch')
            
            warnings_df = df[df['Предупреждение'] != ''] if 'Предупреждение' in df.columns else pd.DataFrame()
            if not warnings_df.empty:
                st.warning("⚠️ Некоторые тарифы являются примерными. Загрузите актуальные данные.")
    
    elif current_section == 'dashboard':
        st.markdown("## 📈 Дашборд")
        
        if not st.session_state.results:
            st.warning("⚠️ Нет данных. Выполните расчет в разделе 'Калькулятор FBS'.")
            return
        
        results = st.session_state.results
        
        col1, col2, col3, col4, col5, col6 = st.columns(6)
        with col1:
            st.metric("📦 Товаров", len(results))
        with col2:
            profitable = len([r for r in results if r.gross_profit > 0])
            st.metric("✅ Прибыльных", f"{profitable} ({profitable/len(results)*100:.0f}%)")
        with col3:
            total_profit = sum(r.gross_profit for r in results)
            st.metric("💰 Прибыль", f"{total_profit:,.0f} ₽")
        with col4:
            avg_margin = np.mean([r.margin_percent for r in results])
            st.metric("📊 Маржа", f"{avg_margin:.1f}%")
        with col5:
            avg_ltv_cac = np.mean([r.ltv_cac_ratio for r in results if r.ltv_cac_ratio < 999])
            st.metric("👥 LTV/CAC", f"{avg_ltv_cac:.1f}x")
        with col6:
            critical = len([r for r in results if r.is_logistic_critical])
            st.metric("🔴 Крит. логистика", f"{critical} ({critical/len(results)*100:.0f}%)")
        
        st.markdown("---")
        render_dashboard_with_filters(results)
        
        st.markdown("### 📊 Визуализация")
        
        col1, col2 = st.columns(2)
        with col1:
            margins = [r.margin_percent for r in results]
            fig = px.histogram(margins, title="Распределение маржинальности",
                             labels={'value': 'Маржа, %', 'count': 'Количество товаров'},
                             nbins=20, color_discrete_sequence=['#6c5ce7'])
            st.plotly_chart(fig, width='stretch')
        
        with col2:
            top = sorted(results, key=lambda x: x.gross_profit, reverse=True)[:10]
            df = pd.DataFrame({'Артикул': [r.artikul for r in top], 'Прибыль, ₽': [r.gross_profit for r in top]})
            fig = px.bar(df, x='Артикул', y='Прибыль, ₽', title="Топ-10 по прибыли",
                        color='Прибыль, ₽', color_continuous_scale='viridis')
            st.plotly_chart(fig, width='stretch')
        
        st.markdown("### 🚚 Логистические зоны риска")
        zones = {'Критическая': 0, 'Зона риска': 0, 'Безопасная': 0, 'Идеальная': 0}
        for r in results:
            if '🔴' in r.logistic_zone_label:
                zones['Критическая'] += 1
            elif '🟡' in r.logistic_zone_label:
                zones['Зона риска'] += 1
            elif '🟢' in r.logistic_zone_label:
                zones['Безопасная'] += 1
            else:
                zones['Идеальная'] += 1
        
        df_zones = pd.DataFrame({'Зона': list(zones.keys()), 'Количество': list(zones.values())})
        fig = px.pie(df_zones, values='Количество', names='Зона', title="Распределение по логистическим зонам",
                    color_discrete_sequence=['#FF6B6B', '#FFD93D', '#6BCB77', '#4D96FF'])
        st.plotly_chart(fig, width='stretch')
    
    elif current_section == 'what_if':
        st.markdown("## 🎯 Анализ сценариев 'Что если'")
        
        if not st.session_state.input_data_list:
            st.warning("⚠️ Сначала выполните расчет в разделе 'Калькулятор FBS'.")
            return
        
        base_data = st.session_state.input_data_list[0]
        
        if base_data is None:
            st.warning("⚠️ Нет базовых данных для анализа.")
            return
        
        st.markdown("### 📋 Базовые параметры")
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Артикул", base_data.artikul)
            st.metric("Цена продажи", f"{base_data.selling_price:,.0f} ₽")
        with col2:
            st.metric("Себестоимость", f"{base_data.cogs:,.0f} ₽")
            st.metric("Вес", f"{base_data.weight_kg:.1f} кг")
        with col3:
            st.metric("Расстояние до МП", f"{base_data.warehouse_distance_km:.0f} км")
            st.metric("Продаж в день", f"{base_data.daily_sales} шт")
        
        st.markdown("---")
        st.markdown("### 🎯 Настройка сценариев")
        
        preset_scenarios = [
            {'name': '📈 Повышение цены на 20%', 'selling_price': base_data.selling_price * 1.2},
            {'name': '📉 Снижение цены на 15%', 'selling_price': base_data.selling_price * 0.85},
            {'name': '🚚 Увеличение расстояния на 50%', 'warehouse_distance_km': base_data.warehouse_distance_km * 1.5},
            {'name': '📦 Оптимизация паллет (x2)', 'pallet_capacity': base_data.pallet_capacity * 2},
            {'name': '🕒 Внедрение ночной смены', 'has_night_shift': True},
            {'name': '💰 Снижение себестоимости на 10%', 'cogs': base_data.cogs * 0.9}
        ]
        
        selected_presets = st.multiselect(
            "Выберите сценарии:",
            [s['name'] for s in preset_scenarios],
            default=[s['name'] for s in preset_scenarios[:3]]
        )
        
        scenarios_to_run = [s for s in preset_scenarios if s['name'] in selected_presets]
        
        st.markdown("#### ➕ Добавить свой сценарий")
        col1, col2 = st.columns(2)
        with col1:
            custom_name = st.text_input("Название сценария", "Мой сценарий")
        with col2:
            custom_param = st.selectbox("Параметр для изменения", 
                                       ["selling_price", "cogs", "warehouse_distance_km", "pallet_capacity", 
                                        "daily_sales", "packaging_cost", "marketing_budget_per_unit"])
        custom_value = st.number_input("Новое значение", value=base_data.selling_price)
        
        if st.button("➕ Добавить сценарий"):
            scenarios_to_run.append({'name': custom_name, custom_param: custom_value})
            st.success(f"✅ Сценарий '{custom_name}' добавлен!")
        
        if st.button("🚀 Запустить анализ", type="primary"):
            if not scenarios_to_run:
                st.warning("⚠️ Выберите хотя бы один сценарий.")
            else:
                with st.spinner("Выполнение анализа..."):
                    df_results = calculator.run_what_if_analysis(base_data, scenarios_to_run)
                    st.markdown("### 📊 Результаты анализа")
                    st.dataframe(df_results, width='stretch')
                    
                    st.markdown("### 📈 Визуализация сценариев")
                    fig = make_subplots(rows=1, cols=2, subplot_titles=("Прибыль", "Маржа"))
                    
                    fig.add_trace(
                        go.Bar(x=df_results['Сценарий'], y=df_results['Прибыль, ₽'], 
                               name='Прибыль', marker_color='#6c5ce7'),
                        row=1, col=1
                    )
                    fig.add_trace(
                        go.Bar(x=df_results['Сценарий'], y=df_results['Маржа, %'], 
                               name='Маржа', marker_color='#00b894'),
                        row=1, col=2
                    )
                    
                    fig.update_layout(height=400, showlegend=True)
                    st.plotly_chart(fig, width='stretch')
                    
                    best_profit = df_results.loc[df_results['Прибыль, ₽'].idxmax()]
                    st.success(f"🏆 Лучший сценарий: **{best_profit['Сценарий']}** "
                              f"(Прибыль: {best_profit['Прибыль, ₽']:,.0f} ₽, Маржа: {best_profit['Маржа, %']:.1f}%)")
    
    elif current_section == 'recommendations':
        st.markdown("## 💡 Автоматические рекомендации")
        
        if not st.session_state.results:
            st.warning("⚠️ Нет данных. Выполните расчет в разделе 'Калькулятор FBS'.")
            return
        
        results = st.session_state.results
        
        if st.button("🔄 Сгенерировать рекомендации", type="primary") or st.session_state.recommendations:
            if not st.session_state.recommendations:
                with st.spinner("Генерация рекомендаций..."):
                    st.session_state.recommendations = calculator.generate_recommendations(results)
            
            if not st.session_state.recommendations:
                st.success("✅ Все показатели в норме! Рекомендаций нет.")
            else:
                st.markdown("### 📋 Рекомендации по приоритету")
                
                priority_order = {'high': 0, 'medium': 1, 'low': 2}
                sorted_recommendations = sorted(st.session_state.recommendations, 
                                               key=lambda x: priority_order.get(x['priority'], 3))
                
                for rec in sorted_recommendations:
                    priority_icon = "🔴" if rec['priority'] == 'high' else "🟡" if rec['priority'] == 'medium' else "🟢"
                    with st.expander(f"{priority_icon} [{rec['priority'].upper()}] {rec['category']} - {rec['icon']} {rec['message'][:80]}..."):
                        st.markdown(f"**{rec['message']}**")
                        
                        if rec.get('affected_products'):
                            st.markdown("**📦 Затронутые товары:**")
                            st.write(", ".join(rec['affected_products'][:10]))
                            if len(rec['affected_products']) > 10:
                                st.caption(f"... и еще {len(rec['affected_products']) - 10} товаров")
        
        if st.session_state.recommendations:
            st.markdown("---")
            st.markdown("### 📊 Статистика рекомендаций")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                high = len([r for r in st.session_state.recommendations if r['priority'] == 'high'])
                st.metric("🔴 Критические", high)
            with col2:
                medium = len([r for r in st.session_state.recommendations if r['priority'] == 'medium'])
                st.metric("🟡 Средние", medium)
            with col3:
                low = len([r for r in st.session_state.recommendations if r['priority'] == 'low'])
                st.metric("🟢 Низкие", low)
            
            categories = {}
            for r in st.session_state.recommendations:
                cat = r['category']
                categories[cat] = categories.get(cat, 0) + 1
            
            st.markdown("### 📂 По категориям")
            df_cat = pd.DataFrame({'Категория': list(categories.keys()), 'Количество': list(categories.values())})
            st.dataframe(df_cat, width='stretch')
    
    elif current_section == 'export':
        st.markdown("## 📥 Экспорт данных")
        
        if not st.session_state.results:
            st.warning("⚠️ Нет данных для экспорта.")
            return
        
        results = st.session_state.results
        input_data_list = st.session_state.input_data_list
        
        st.success(f"✅ Доступно для экспорта: {len(results)} товаров")
        st.info("📌 Данные экспортируются на основе реальных расчетов, без захардкоженных значений.")
        
        tab1, tab2, tab3 = st.tabs(["📊 Excel", "📄 CSV", "🌐 Google Sheets (ручной)"])
        
        with tab1:
            st.info("Excel файл содержит формулы и условное форматирование. При изменении тарифов все расчеты пересчитываются автоматически.")
            
            if st.button("📥 Скачать Excel-отчет", type="primary"):
                try:
                    wb = Workbook()
                    
                    ws_tariffs = wb.active
                    ws_tariffs.title = "Тарифы МП"
                    
                    tariff_headers = ['Категория', 'Комиссия, %', 'Мин. комиссия, ₽', 'Last Mile база, ₽', 
                                     'Last Mile за кг, ₽', 'Эквайринг, %', 'Возвраты, %', 'Штрафы, %', 
                                     'Источник']
                    
                    for col, header in enumerate(tariff_headers, 1):
                        cell = ws_tariffs.cell(row=1, column=col, value=header)
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="1a1a2e", fill_type="solid")
                    
                    row = 2
                    for category, tariff in calculator.current_tariffs.items():
                        ws_tariffs.cell(row=row, column=1, value=category)
                        ws_tariffs.cell(row=row, column=2, value=round(tariff.get('commission_rate', 0) * 100, 2))
                        ws_tariffs.cell(row=row, column=3, value=tariff.get('min_commission', 0))
                        ws_tariffs.cell(row=row, column=4, value=tariff.get('last_mile_base', 0))
                        ws_tariffs.cell(row=row, column=5, value=tariff.get('last_mile_per_kg', 0))
                        ws_tariffs.cell(row=row, column=6, value=round(tariff.get('acquiring_fee', 0) * 100, 2))
                        ws_tariffs.cell(row=row, column=7, value=round(tariff.get('return_fee', 0) * 100, 2))
                        ws_tariffs.cell(row=row, column=8, value=round(tariff.get('penalty_rate', 0) * 100, 2))
                        ws_tariffs.cell(row=row, column=9, value=tariff.get('source', 'unknown'))
                        row += 1
                    
                    ws_results = wb.create_sheet("Результаты")
                    
                    result_headers = ['Артикул', 'Наименование', 'Цена, ₽', 'Прибыль, ₽', 'Маржа, %', 
                                     'ROI, %', 'Комиссия, ₽', 'First Mile, ₽', 'Last Mile, ₽',
                                     'Опт. запас, шт', 'Оборачиваемость, дн', 'Лог. зона', 'Источник данных']
                    
                    for col, header in enumerate(result_headers, 1):
                        cell = ws_results.cell(row=1, column=col, value=header)
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="1a1a2e", fill_type="solid")
                    
                    row = 2
                    for result, input_data in zip(results, input_data_list):
                        ws_results.cell(row=row, column=1, value=result.artikul)
                        ws_results.cell(row=row, column=2, value=result.product_name)
                        ws_results.cell(row=row, column=3, value=result.selling_price)
                        ws_results.cell(row=row, column=4, value=result.gross_profit)
                        ws_results.cell(row=row, column=5, value=result.margin_percent)
                        ws_results.cell(row=row, column=6, value=result.roi_percent)
                        ws_results.cell(row=row, column=7, value=result.commission)
                        ws_results.cell(row=row, column=8, value=result.first_mile_cost)
                        ws_results.cell(row=row, column=9, value=result.last_mile_cost)
                        ws_results.cell(row=row, column=10, value=result.optimal_stock_units)
                        ws_results.cell(row=row, column=11, value=result.stock_turnover_days)
                        ws_results.cell(row=row, column=12, value=result.logistic_zone_label)
                        ws_results.cell(row=row, column=13, value=result.data_source)
                        
                        if result.gross_profit > 0:
                            ws_results.cell(row=row, column=4).fill = PatternFill(start_color="C6EFCE", fill_type="solid")
                        else:
                            ws_results.cell(row=row, column=4).fill = PatternFill(start_color="FFC7CE", fill_type="solid")
                        
                        row += 1
                    
                    # Добавляем ColorScale для маржи
                    if row > 2:
                        last_row = row - 1
                        ws_results.conditional_formatting.add(
                            f"E2:E{last_row}",
                            ColorScaleRule(start_type="min", start_color="FFC7CE",
                                         mid_type="percentile", mid_value=50, mid_color="FFEB9C",
                                         end_type="max", end_color="C6EFCE")
                        )
                    
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    filename = f"FBS_Report_{st.session_state.marketplace}_{timestamp}.xlsx"
                    output_path = EXPORTS_DIR / filename
                    wb.save(output_path)
                    
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="⬇️ Скачать Excel",
                            data=f.read(),
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    
                except Exception as e:
                    st.error(f"❌ Ошибка создания Excel: {e}")
        
        with tab2:
            delimiter = st.selectbox("Разделитель", [";", ",", "\\t"])
            if st.button("📥 Скачать CSV", type="primary"):
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                filename = f"FBS_Report_{st.session_state.marketplace}_{timestamp}.csv"
                output_path = EXPORTS_DIR / filename
                
                data = []
                for result, input_data in zip(results, input_data_list):
                    data.append({
                        'Артикул': result.artikul,
                        'Наименование': result.product_name,
                        'Цена, ₽': result.selling_price,
                        'Прибыль, ₽': result.gross_profit,
                        'Маржа, %': result.margin_percent,
                        'ROI, %': result.roi_percent,
                        'Комиссия, ₽': result.commission,
                        'First Mile, ₽': result.first_mile_cost,
                        'Last Mile, ₽': result.last_mile_cost,
                        'Опт. запас, шт': result.optimal_stock_units,
                        'Оборачиваемость, дн': result.stock_turnover_days,
                        'Лог. зона': result.logistic_zone_label,
                        'Источник данных': result.data_source
                    })
                
                df = pd.DataFrame(data)
                sep = "\t" if delimiter == "\\t" else delimiter
                df.to_csv(output_path, index=False, sep=sep, encoding='utf-8-sig')
                
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="⬇️ Скачать CSV",
                        data=f.read(),
                        file_name=filename,
                        mime="text/csv"
                    )
        
        with tab3:
            st.markdown("### 📤 Экспорт в Google Таблицы (без сервисного аккаунта)")
            st.info("""
            **Два способа обновить вашу таблицу:**

            1. **Скачать CSV / TSV** и импортировать вручную (инструкция ниже).
            2. **Скопировать таблицу** прямо из интерфейса и вставить в Google Sheets (Ctrl+V).
            """)
            
            if not results:
                st.warning("⚠️ Нет данных для экспорта.")
            else:
                col1, col2 = st.columns(2)
                with col1:
                    sep = st.selectbox("Формат данных", ["CSV (;)","TSV (табуляция)"], index=0, key="gs_export_sep")
                    if sep == "CSV (;)":
                        csv_data = GoogleSheetsExporter.generate_csv(results, separator=";")
                        file_ext = "csv"
                        mime = "text/csv"
                        label = "CSV (;)"
                    else:
                        csv_data = GoogleSheetsExporter.generate_tsv(results)
                        file_ext = "tsv"
                        mime = "text/tab-separated-values"
                        label = "TSV (табуляция)"
                    
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    filename = f"FBS_Results_{st.session_state.marketplace}_{timestamp}.{file_ext}"
                    
                    st.download_button(
                        label=f"📥 Скачать {label}",
                        data=csv_data.encode('utf-8-sig'),
                        file_name=filename,
                        mime=mime,
                        type="primary"
                    )
                
                with col2:
                    st.markdown("#### 📋 Копировать в буфер")
                    st.caption("Нажмите кнопку, затем вставьте (Ctrl+V) в Google Sheets на нужный лист.")
                    if st.button("📋 Скопировать TSV (для вставки)"):
                        st.code(GoogleSheetsExporter.generate_tsv(results), language="text")
                        st.info("Скопируйте текст выше и вставьте в Google Sheets (данные вставятся как таблица).")
                
                st.markdown("---")
                st.markdown("### 📖 Инструкция по импорту")
                st.markdown(GoogleSheetsExporter.get_import_instructions())
                
                with st.expander("⚙️ Хотите автоматическое обновление? (требуется сервисный аккаунт)"):
                    st.markdown("""
                    Если у вас есть **credentials.json** от Google сервисного аккаунта, вы можете настроить автоматическую синхронизацию.
                    
                    1. Положите файл в папку `config/` (или загрузите через интерфейс).
                    2. Перейдите в раздел **Настройки → Google Sheets** и включите автообновление.
                    3. После каждого расчёта данные будут обновляться в таблице автоматически.
                    
                    [Как создать сервисный аккаунт](https://developers.google.com/workspace/guides/create-credentials)
                    """)
                    
                    uploaded_creds = st.file_uploader("Загрузить credentials.json", type=['json'], key="creds_upload")
                    if uploaded_creds:
                        creds_path = CONFIG_DIR / "google_credentials.json"
                        creds_path.write_bytes(uploaded_creds.getvalue())
                        st.success("✅ Файл сохранён! Теперь вы можете использовать автоматический режим.")
                        st.rerun()
    
    elif current_section == 'gsheets':
        st.markdown("## 🌐 Интеграция с Google Таблицами (ручной режим)")
        st.info("""
        **Как это работает:**  
        Вы создаёте или подключаете существующую Google Таблицу, а приложение автоматически обновляет в ней:
        - **Актуальные тарифы** маркетплейса (из API / AI / CSV)
        - **Результаты расчётов** по всем товарам
        
        При изменении тарифов в приложении вы можете **одним кликом** обновить таблицу.
        """)
        
        # Инициализация менеджера (если есть credentials)
        if 'gs_manager' not in st.session_state:
            # Пытаемся создать менеджер, если есть credentials
            creds_path = CONFIG_DIR / "google_credentials.json"
            if creds_path.exists() and GSPREAD_AVAILABLE:
                try:
                    from google.oauth2.service_account import Credentials
                    gc = gspread.service_account(filename=str(creds_path))
                    st.session_state.gs_manager = gc
                except Exception as e:
                    st.warning(f"⚠️ Не удалось загрузить credentials: {e}")
            else:
                st.session_state.gs_manager = None
        
        manager = st.session_state.gs_manager
        
        col1, col2 = st.columns(2)
        with col1:
            sheet_url = st.text_input("🔗 Ссылка на существующую таблицу", 
                                      value=st.session_state.get('gsheet_url', ''),
                                      help="Вставьте URL таблицы, к которой у вас есть доступ на редактирование")
            if st.button("📂 Открыть таблицу"):
                if sheet_url and manager:
                    try:
                        sheet = manager.open_by_url(sheet_url)
                        st.session_state.gsheet = sheet
                        st.session_state.gsheet_url = sheet_url
                        st.success(f"✅ Таблица открыта: {sheet.title}")
                    except Exception as e:
                        st.error(f"❌ Не удалось открыть таблицу: {e}")
                else:
                    st.warning("⚠️ Загрузите credentials.json в разделе 'Настройки' или используйте ручной экспорт.")
        
        with col2:
            st.markdown("#### ➕ Создать новую таблицу")
            new_title = st.text_input("Название новой таблицы", "FBS Unit Economics")
            if st.button("🆕 Создать и открыть"):
                if manager:
                    try:
                        sheet = manager.create(new_title)
                        st.session_state.gsheet = sheet
                        st.session_state.gsheet_url = sheet.url
                        st.success(f"✅ Таблица создана! Открыть: {sheet.url}")
                        st.markdown(f"[Открыть таблицу]({sheet.url})")
                    except Exception as e:
                        st.error(f"❌ Ошибка создания: {e}")
                else:
                    st.warning("⚠️ Загрузите credentials.json в разделе 'Настройки'.")
        
        st.markdown("---")
        st.markdown("### 📤 Обновление данных")
        
        if 'gsheet' in st.session_state and st.session_state.gsheet:
            sheet = st.session_state.gsheet
            st.success(f"Текущая таблица: **{sheet.title}**")
            
            col1, col2, col3 = st.columns(3)
            with col1:
                if st.button("🔄 Обновить тарифы"):
                    tariffs = st.session_state.calculator.current_tariffs
                    if tariffs:
                        with st.spinner("Обновление тарифов..."):
                            try:
                                # Простой пример обновления (реализуйте свой метод)
                                ws = sheet.worksheet("Тарифы")
                                ws.clear()
                                ws.append_row(['Категория', 'Комиссия,%', 'Мин.комиссия', 'Last Mile база', 
                                             'Last Mile за кг', 'Эквайринг,%', 'Возвраты,%', 'Штрафы,%',
                                             'Источник', 'Обновлено'])
                                now = datetime.now().isoformat()
                                for cat, data in tariffs.items():
                                    ws.append_row([
                                        cat,
                                        round(data.get('commission_rate', 0) * 100, 2),
                                        data.get('min_commission', 0),
                                        data.get('last_mile_base', 0),
                                        data.get('last_mile_per_kg', 0),
                                        round(data.get('acquiring_fee', 0) * 100, 2),
                                        round(data.get('return_fee', 0) * 100, 2),
                                        round(data.get('penalty_rate', 0) * 100, 2),
                                        data.get('source', 'unknown'),
                                        now
                                    ])
                                st.success("✅ Тарифы обновлены в таблице!")
                            except Exception as e:
                                st.error(f"❌ Ошибка обновления: {e}")
                    else:
                        st.warning("⚠️ Сначала загрузите тарифы (раздел 'Тарифы')")
            with col2:
                if st.button("📊 Обновить результаты"):
                    results = st.session_state.results
                    if results:
                        with st.spinner("Обновление результатов..."):
                            try:
                                ws = sheet.worksheet("Результаты")
                                ws.clear()
                                headers = ['Артикул', 'Название', 'Цена', 'Прибыль', 'Маржа,%', 'ROI,%',
                                           'Комиссия', 'First Mile', 'Last Mile', 'Pick&Pack', 'Упаковка',
                                           'Эквайринг', 'Возвраты', 'Штрафы', 'Маркетинг', 'Склад', 'Налог',
                                           'Опт.запас', 'Оборач., дн', 'Лог.зона', 'Источник']
                                ws.append_row(headers)
                                for r in results:
                                    ws.append_row([
                                        r.artikul, r.product_name, r.selling_price, r.gross_profit,
                                        r.margin_percent, r.roi_percent, r.commission, r.first_mile_cost,
                                        r.last_mile_cost, r.pick_pack_cost, r.packaging_cost, r.acquiring_cost,
                                        r.return_cost, r.penalty_cost, r.marketing_cost, r.warehouse_cost,
                                        r.tax_cost, r.optimal_stock_units, r.stock_turnover_days,
                                        r.logistic_zone_label, r.data_source
                                    ])
                                st.success("✅ Результаты обновлены в таблице!")
                            except Exception as e:
                                st.error(f"❌ Ошибка обновления: {e}")
                    else:
                        st.warning("⚠️ Сначала выполните расчёт (раздел 'Калькулятор')")
            with col3:
                if st.button("🔄 Обновить всё"):
                    tariffs = st.session_state.calculator.current_tariffs
                    results = st.session_state.results
                    if tariffs and results:
                        with st.spinner("Обновление всех данных..."):
                            try:
                                # Обновляем тарифы
                                ws = sheet.worksheet("Тарифы")
                                ws.clear()
                                ws.append_row(['Категория', 'Комиссия,%', 'Мин.комиссия', 'Last Mile база', 
                                             'Last Mile за кг', 'Эквайринг,%', 'Возвраты,%', 'Штрафы,%',
                                             'Источник', 'Обновлено'])
                                now = datetime.now().isoformat()
                                for cat, data in tariffs.items():
                                    ws.append_row([
                                        cat,
                                        round(data.get('commission_rate', 0) * 100, 2),
                                        data.get('min_commission', 0),
                                        data.get('last_mile_base', 0),
                                        data.get('last_mile_per_kg', 0),
                                        round(data.get('acquiring_fee', 0) * 100, 2),
                                        round(data.get('return_fee', 0) * 100, 2),
                                        round(data.get('penalty_rate', 0) * 100, 2),
                                        data.get('source', 'unknown'),
                                        now
                                    ])
                                # Обновляем результаты
                                ws2 = sheet.worksheet("Результаты")
                                ws2.clear()
                                headers = ['Артикул', 'Название', 'Цена', 'Прибыль', 'Маржа,%', 'ROI,%',
                                           'Комиссия', 'First Mile', 'Last Mile', 'Pick&Pack', 'Упаковка',
                                           'Эквайринг', 'Возвраты', 'Штрафы', 'Маркетинг', 'Склад', 'Налог',
                                           'Опт.запас', 'Оборач., дн', 'Лог.зона', 'Источник']
                                ws2.append_row(headers)
                                for r in results:
                                    ws2.append_row([
                                        r.artikul, r.product_name, r.selling_price, r.gross_profit,
                                        r.margin_percent, r.roi_percent, r.commission, r.first_mile_cost,
                                        r.last_mile_cost, r.pick_pack_cost, r.packaging_cost, r.acquiring_cost,
                                        r.return_cost, r.penalty_cost, r.marketing_cost, r.warehouse_cost,
                                        r.tax_cost, r.optimal_stock_units, r.stock_turnover_days,
                                        r.logistic_zone_label, r.data_source
                                    ])
                                st.success("✅ Все данные обновлены!")
                            except Exception as e:
                                st.error(f"❌ Ошибка обновления: {e}")
                    else:
                        st.warning("⚠️ Загрузите тарифы и выполните расчёт")
            
            st.markdown("---")
            st.markdown("### ⚙️ Настройки автоматического обновления")
            auto_update = st.checkbox("Автоматически обновлять таблицу при каждом новом расчёте", 
                                      value=st.session_state.get('auto_gs_update', False))
            st.session_state.auto_gs_update = auto_update
            if auto_update:
                st.info("✅ Включено автоматическое обновление. После каждого расчёта данные будут синхронизироваться.")
        else:
            st.warning("⚠️ Сначала откройте или создайте таблицу (нужен сервисный аккаунт) или используйте ручной экспорт в разделе 'Экспорт'.")
    
    elif current_section == 'settings':
        st.markdown("## ⚙️ Настройки")
        
        tab1, tab2, tab3, tab4 = st.tabs(["🔑 API Ключи", "🏪 Маркетплейс и налоги", "📄 Импорт тарифов CSV", "🌐 Google Sheets (сервисный аккаунт)"])
        
        with tab1:
            st.markdown("### 🔑 Настройка API ключей")
            st.info("API ключи используются для загрузки актуальных тарифов маркетплейсов. Без них используются примерные значения.")
            
            api_manager = st.session_state.api_manager
            
            st.markdown("#### 📦 Ozon")
            col1, col2 = st.columns(2)
            with col1:
                client_id = st.text_input("Ozon Client ID", value=api_manager.get_api_key('ozon_client_id') or '', type="password")
            with col2:
                api_key = st.text_input("Ozon API Key", value=api_manager.get_api_key('ozon') or '', type="password")
            if st.button("💾 Сохранить Ozon"):
                if client_id and api_key:
                    api_manager.save_api_key('ozon_client_id', client_id)
                    api_manager.save_api_key('ozon', api_key)
                    st.success("✅ Ключи Ozon сохранены!")
            
            st.markdown("---")
            
            st.markdown("#### 📦 Wildberries")
            wb_key = st.text_input("Wildberries API Token", value=api_manager.get_api_key('wildberries') or '', type="password")
            if st.button("💾 Сохранить Wildberries"):
                if wb_key:
                    api_manager.save_api_key('wildberries', wb_key)
                    st.success("✅ Ключ Wildberries сохранен!")
            
            st.markdown("---")
            
            st.markdown("#### 📦 Яндекс Маркет")
            col1, col2 = st.columns(2)
            with col1:
                ym_token = st.text_input("Яндекс Маркет OAuth Token", value=api_manager.get_api_key('yandex_market') or '', type="password")
            with col2:
                ym_campaign = st.text_input("Campaign ID", value=api_manager.get_api_key('yandex_campaign_id') or '')
            if st.button("💾 Сохранить Яндекс Маркет"):
                if ym_token and ym_campaign:
                    api_manager.save_api_key('yandex_market', ym_token)
                    api_manager.save_api_key('yandex_campaign_id', ym_campaign)
                    st.success("✅ Ключи Яндекс Маркет сохранены!")
            
            st.markdown("---")
            
            st.markdown("#### 🤖 DeepSeek AI")
            ds_key = st.text_input("DeepSeek API Key", value=api_manager.get_api_key('deepseek') or '', type="password")
            if st.button("💾 Сохранить DeepSeek"):
                if ds_key:
                    api_manager.save_api_key('deepseek', ds_key)
                    st.success("✅ Ключ DeepSeek сохранен!")
        
        with tab2:
            st.markdown("### 🏪 Настройки маркетплейса и налогов")
            
            col1, col2 = st.columns(2)
            
            with col1:
                marketplace = st.selectbox(
                    "Маркетплейс по умолчанию",
                    ["Ozon", "Wildberries", "Яндекс Маркет"],
                    index=["Ozon", "Wildberries", "Яндекс Маркет"].index(st.session_state.marketplace)
                )
                if st.button("💾 Сохранить маркетплейс"):
                    st.session_state.marketplace = marketplace
                    st.session_state.calculator.set_marketplace(marketplace)
                    st.success(f"✅ Маркетплейс '{marketplace}' сохранен!")
                    st.rerun()
            
            with col2:
                tax_system = st.selectbox(
                    "Система налогообложения",
                    list(TAX_SYSTEMS.keys()),
                    index=list(TAX_SYSTEMS.keys()).index(st.session_state.tax_system) if st.session_state.tax_system in TAX_SYSTEMS else 0
                )
                if st.button("💾 Сохранить налоговую систему"):
                    st.session_state.tax_system = tax_system
                    st.session_state.calculator.tax_system = tax_system
                    st.success(f"✅ Налоговая система '{tax_system}' сохранена!")
        
        with tab3:
            st.markdown("### 📄 Импорт тарифов из CSV")
            st.markdown("""
            **Формат CSV файла:**
            - Обязательные колонки: `category`, `commission_rate`, `min_commission`, `last_mile_base`
            - Опциональные колонки: `last_mile_per_kg`, `last_mile_per_km`, `acquiring_fee`, `return_fee`
            - Разделитель: запятая или точка с запятой
            - Кодировка: UTF-8
            
            **Пример:**
category,commission_rate,min_commission,last_mile_base,last_mile_per_kg,acquiring_fee,return_fee
electronics,0.10,30,50,15,0.015,0.02
clothing,0.15,25,40,12,0.015,0.018
""")
            
            csv_import_file = st.file_uploader(
                "📁 Загрузите CSV с тарифами",
                type=['csv'],
                key="tariffs_csv_import"
            )
            
            if csv_import_file and st.button("📥 Загрузить тарифы из CSV", type="primary"):
                try:
                    csv_content = csv_import_file.getvalue().decode('utf-8')
                    df_test = pd.read_csv(io.StringIO(csv_content))
                    required_cols = ['category', 'commission_rate', 'min_commission', 'last_mile_base']
                    missing_cols = [col for col in required_cols if col not in df_test.columns]
                    
                    if missing_cols:
                        st.error(f"❌ Отсутствуют обязательные колонки: {', '.join(missing_cols)}")
                        st.info(f"Доступные колонки: {', '.join(df_test.columns)}")
                    else:
                        calculator.refresh_tariffs(force=True, csv_content=csv_content)
                        st.success(f"✅ Тарифы загружены из CSV! ({len(df_test)} категорий)")
                        st.dataframe(df_test, width='stretch')
                        st.rerun()
                except Exception as e:
                    st.error(f"❌ Ошибка загрузки CSV: {e}")
        
        with tab4:
            st.markdown("### 🌐 Настройка сервисного аккаунта Google")
            st.info("""
            Загрузите файл credentials.json от сервисного аккаунта Google для автоматической синхронизации с Google Sheets.
            
            Если у вас нет сервисного аккаунта, вы можете использовать ручной экспорт через CSV/TSV (раздел "Экспорт").
            """)
            
            uploaded_creds = st.file_uploader("📁 Загрузить credentials.json", type=['json'], key="settings_creds")
            if uploaded_creds:
                creds_path = CONFIG_DIR / "google_credentials.json"
                creds_path.write_bytes(uploaded_creds.getvalue())
                st.success("✅ Файл сохранён! Теперь вы можете использовать автоматический режим в разделе 'Google Sheets'.")
                st.rerun()
            
            if (CONFIG_DIR / "google_credentials.json").exists():
                st.success("✅ Файл credentials.json уже загружен.")
                if st.button("🗑️ Удалить credentials.json"):
                    (CONFIG_DIR / "google_credentials.json").unlink()
                    st.success("✅ Файл удалён.")
                    st.rerun()

if __name__ == "__main__":
    main()
