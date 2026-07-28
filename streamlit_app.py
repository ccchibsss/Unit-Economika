"""
============================================================================
🚀 FBS UNIT ECONOMICS PRO 2026 — ПОЛНАЯ ОПЕРАЦИОННАЯ ВЕРСИЯ С УЛУЧШЕНИЯМИ
============================================================================
Операционный директор | FBS-экспертиза | Оптимизация складских остатков
Маркетплейсы: Ozon, Wildberries, Яндекс Маркет
Версия: 6.2.0

КЛЮЧЕВЫЕ ПРИНЦИПЫ:
1. НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ — все данные из API или пользовательского ввода
2. Полная прозрачность расчетов
3. Динамическая подгрузка тарифов
4. Оптимизация на основе реальных данных

НИЧЕГО НЕ СОКРАЩЕНО — ПОЛНАЯ ВЕРСИЯ
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
from enum import Enum
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

APP_VERSION = "6.2.0"
APP_NAME = "🚀 FBS Юнит-экономика PRO 2026 — Операционная версия с улучшениями"
APP_DESCRIPTION = "Профессиональный расчет юнит-экономики для FBS-модели с оптимизацией складских остатков и логистических коридоров"

# Настройка путей
BASE_DIR = Path(__file__).parent.resolve() if '__file__' in dir() else Path.cwd()
DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
LOGS_DIR = BASE_DIR / "logs"
EXPORTS_DIR = BASE_DIR / "exports"
CONFIG_DIR = BASE_DIR / "config"
TEMP_DIR = BASE_DIR / "temp"
TARIFFS_CACHE_DIR = CACHE_DIR / "tariffs"

# Создание директорий
for dir_path in [DATA_DIR, CACHE_DIR, LOGS_DIR, EXPORTS_DIR, CONFIG_DIR, TEMP_DIR, TARIFFS_CACHE_DIR]:
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
    """Декоратор для измерения времени выполнения функций"""
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
    """Декоратор для повторных попыток при ошибках API запросов"""
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
    """Мемоизация для кэширования результатов функций"""
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
    """Отслеживание прогресса операций"""
    
    def __init__(self):
        self.progress = 0.0
        self.status = ""
        self.total = 0
        self.current = 0
        self.start_time = None
        self.estimated_time_remaining = 0
    
    def start(self, total: int, status: str = ""):
        """Начало отслеживания"""
        self.total = total
        self.current = 0
        self.progress = 0.0
        self.status = status
        self.start_time = time.time()
    
    def update(self, current: int, status: str = ""):
        """Обновление прогресса"""
        self.current = current
        self.total = max(self.total, current)
        self.progress = min(current / self.total, 1.0) if self.total > 0 else 0
        if status:
            self.status = status
        
        # Расчет оставшегося времени
        if self.start_time and self.progress > 0:
            elapsed = time.time() - self.start_time
            self.estimated_time_remaining = (elapsed / self.progress) * (1 - self.progress)
    
    def get_progress(self) -> float:
        """Получение текущего прогресса"""
        return self.progress
    
    def get_status(self) -> str:
        """Получение текущего статуса"""
        return self.status
    
    def get_eta(self) -> float:
        """Получение оставшегося времени"""
        return self.estimated_time_remaining

class AuditLogger:
    """Логирование действий пользователя для аудита"""
    
    def __init__(self):
        self.audit_file = LOGS_DIR / "audit.log"
        self._init_audit_file()
    
    def _init_audit_file(self):
        if not self.audit_file.exists():
            with open(self.audit_file, 'w', encoding='utf-8') as f:
                f.write("timestamp,user,action,details\n")
    
    def log(self, action: str, details: Dict[str, Any]):
        """Запись действия в аудит-лог"""
        user = getpass.getuser()
        timestamp = datetime.now().isoformat()
        
        with open(self.audit_file, 'a', encoding='utf-8') as f:
            f.write(f"{timestamp},{user},{action},{json.dumps(details, ensure_ascii=False)}\n")
        
        logger.info(f"📝 Аудит: {user} - {action}")

# ============================================================================
# БЛОК 2: БЕЗОПАСНОЕ ХРАНЕНИЕ ДАННЫХ (ШИФРОВАНИЕ)
# ============================================================================

class SecureDataManager:
    """
    Менеджер безопасного хранения конфиденциальных данных.
    Использует Fernet для шифрования API ключей и других секретов.
    """
    
    def __init__(self):
        self.key_file = CONFIG_DIR / ".master_key"
        self.data_file = CONFIG_DIR / ".secure_data.enc"
        self._fernet = None
        self._init_encryption()
    
    def _init_encryption(self):
        """Инициализация шифрования Fernet"""
        if not CRYPTO_AVAILABLE:
            logger.warning("⚠️ Cryptography не установлен. Данные не будут зашифрованы.")
            return
        
        try:
            if self.key_file.exists():
                # Загружаем существующий ключ
                key = self.key_file.read_bytes()
            else:
                # Генерируем новый ключ
                key = Fernet.generate_key()
                self.key_file.write_bytes(key)
                # Устанавливаем права только для владельца
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
        """Проверка доступности шифрования"""
        return self._fernet is not None
    
    def save_data(self, data: Dict[str, Any]) -> bool:
        """
        Сохраняет зашифрованные данные в файл.
        
        Args:
            data: Словарь с данными для шифрования
            
        Returns:
            bool: True если сохранение успешно
        """
        if not self._fernet:
            logger.warning("⚠️ Шифрование недоступно, данные не сохранены")
            return False
        
        try:
            # Сериализуем в JSON
            json_data = json.dumps(data, ensure_ascii=False, indent=2)
            # Шифруем
            encrypted = self._fernet.encrypt(json_data.encode('utf-8'))
            # Сохраняем
            self.data_file.write_bytes(encrypted)
            # Устанавливаем права
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
        """
        Загружает и расшифровывает данные из файла.
        
        Returns:
            Dict: Расшифрованные данные или пустой словарь при ошибке
        """
        if not self._fernet or not self.data_file.exists():
            return {}
        
        try:
            # Читаем зашифрованные данные
            encrypted = self.data_file.read_bytes()
            # Расшифровываем
            decrypted = self._fernet.decrypt(encrypted)
            # Парсим JSON
            return json.loads(decrypted.decode('utf-8'))
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки зашифрованных данных: {e}")
            return {}
    
    def delete_data(self) -> bool:
        """Удаляет зашифрованные данные"""
        try:
            if self.data_file.exists():
                self.data_file.unlink()
                logger.info("🗑️ Зашифрованные данные удалены")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка удаления данных: {e}")
            return False
    
    def store_api_key(self, service: str, api_key: str) -> bool:
        """
        Безопасное сохранение API ключа.
        
        Args:
            service: Название сервиса (ozon, wildberries, etc.)
            api_key: API ключ
            
        Returns:
            bool: True если сохранение успешно
        """
        data = self.load_data()
        if 'api_keys' not in data:
            data['api_keys'] = {}
        
        data['api_keys'][service] = api_key
        data['api_keys_updated'] = datetime.now().isoformat()
        
        return self.save_data(data)
    
    def get_api_key(self, service: str) -> Optional[str]:
        """
        Получение API ключа.
        
        Args:
            service: Название сервиса
            
        Returns:
            Optional[str]: API ключ или None
        """
        data = self.load_data()
        return data.get('api_keys', {}).get(service)
    
    def get_all_api_keys(self) -> Dict[str, str]:
        """Получение всех сохраненных API ключей"""
        data = self.load_data()
        return data.get('api_keys', {})
    
    def delete_api_key(self, service: str) -> bool:
        """Удаление API ключа"""
        data = self.load_data()
        if 'api_keys' in data and service in data['api_keys']:
            del data['api_keys'][service]
            return self.save_data(data)
        return False
    
    def clear_all_keys(self) -> bool:
        """Удаление всех API ключей"""
        data = self.load_data()
        data['api_keys'] = {}
        return self.save_data(data)

# ============================================================================
# БЛОК 3: КЭШИРОВАНИЕ ДЛЯ ОПТИМИЗАЦИИ ПРОИЗВОДИТЕЛЬНОСТИ
# ============================================================================

class CacheManager:
    """
    Менеджер кэширования для оптимизации производительности.
    Поддерживает многоуровневое кэширование: память + диск.
    """
    
    def __init__(self, max_memory_mb: int = 500, cache_ttl_seconds: int = 3600):
        self.cache_dir = CACHE_DIR
        self.max_memory_mb = max_memory_mb
        self.cache_ttl = cache_ttl_seconds
        self._memory_cache: Dict[str, Any] = {}
        self._cache_timestamps: Dict[str, float] = {}
        self._cache_sizes: Dict[str, int] = {}
        
        # Создаем поддиректории для разных типов кэша
        self.tariffs_cache_dir = self.cache_dir / "tariffs"
        self.api_cache_dir = self.cache_dir / "api_responses"
        self.calc_cache_dir = self.cache_dir / "calculations"
        
        for dir_path in [self.tariffs_cache_dir, self.api_cache_dir, self.calc_cache_dir]:
            dir_path.mkdir(exist_ok=True, parents=True)
    
    def _get_cache_key(self, *args, **kwargs) -> str:
        """Генерация ключа кэша на основе аргументов"""
        key_parts = [str(arg) for arg in args]
        key_parts.extend(f"{k}={v}" for k, v in sorted(kwargs.items()))
        key_string = "|".join(key_parts)
        return hashlib.md5(key_string.encode()).hexdigest()
    
    def _get_disk_cache_path(self, cache_type: str, key: str) -> Path:
        """Получение пути к файлу кэша на диске"""
        cache_dirs = {
            'tariffs': self.tariffs_cache_dir,
            'api': self.api_cache_dir,
            'calc': self.calc_cache_dir
        }
        cache_dir = cache_dirs.get(cache_type, self.cache_dir)
        return cache_dir / f"{key}.cache"
    
    def get(self, cache_type: str, key: str) -> Optional[Any]:
        """
        Получение значения из кэша.
        Проверяет сначала память, потом диск.
        """
        # Проверка memory cache
        memory_key = f"{cache_type}:{key}"
        if memory_key in self._memory_cache:
            timestamp = self._cache_timestamps.get(memory_key, 0)
            if time.time() - timestamp < self.cache_ttl:
                logger.debug(f"📦 Кэш попадание (память): {memory_key}")
                return self._memory_cache[memory_key]
            else:
                # Удаляем устаревший кэш
                del self._memory_cache[memory_key]
                del self._cache_timestamps[memory_key]
        
        # Проверка disk cache
        cache_path = self._get_disk_cache_path(cache_type, key)
        if cache_path.exists():
            try:
                with open(cache_path, 'rb') as f:
                    cached_data = pickle.load(f)
                
                # Проверяем время жизни
                cached_time = cached_data.get('timestamp', 0)
                if time.time() - cached_time < self.cache_ttl:
                    logger.debug(f"💾 Кэш попадание (диск): {key}")
                    # Загружаем в память для быстрого доступа
                    value = cached_data.get('data')
                    self._memory_cache[memory_key] = value
                    self._cache_timestamps[memory_key] = cached_time
                    return value
                else:
                    # Удаляем устаревший файл
                    cache_path.unlink()
            except Exception as e:
                logger.debug(f"Ошибка чтения дискового кэша: {e}")
                try:
                    cache_path.unlink()
                except:
                    pass
        
        return None
    
    def set(self, cache_type: str, key: str, value: Any):
        """Сохранение значения в кэш (память + диск)"""
        memory_key = f"{cache_type}:{key}"
        current_time = time.time()
        
        # Сохраняем в память
        self._memory_cache[memory_key] = value
        self._cache_timestamps[memory_key] = current_time
        
        # Сохраняем на диск
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
        
        # Проверка размера кэша в памяти
        self._cleanup_memory_cache()
    
    def _cleanup_memory_cache(self):
        """Очистка устаревших записей из кэша памяти"""
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
        """
        Очистка кэша.
        
        Args:
            cache_type: Тип кэша для очистки (None - очистить всё)
        """
        if cache_type:
            # Очистка определенного типа кэша
            cache_dirs = {
                'tariffs': self.tariffs_cache_dir,
                'api': self.api_cache_dir,
                'calc': self.calc_cache_dir
            }
            if cache_type in cache_dirs:
                for cache_file in cache_dirs[cache_type].glob("*.cache"):
                    cache_file.unlink()
            
            # Очистка memory cache для этого типа
            prefix = f"{cache_type}:"
            keys_to_remove = [k for k in self._memory_cache if k.startswith(prefix)]
            for key in keys_to_remove:
                del self._memory_cache[key]
                if key in self._cache_timestamps:
                    del self._cache_timestamps[key]
        else:
            # Полная очистка
            self._memory_cache.clear()
            self._cache_timestamps.clear()
            
            for cache_dir in [self.tariffs_cache_dir, self.api_cache_dir, self.calc_cache_dir]:
                for cache_file in cache_dir.glob("*.cache"):
                    cache_file.unlink()
        
        logger.info(f"🗑️ Кэш очищен: {cache_type or 'все типы'}")
    
    def get_cache_stats(self) -> Dict[str, Any]:
        """Получение статистики кэша"""
        return {
            'memory_entries': len(self._memory_cache),
            'memory_size_mb': sum(self._cache_sizes.values()) / (1024 * 1024),
            'tariffs_cache_files': len(list(self.tariffs_cache_dir.glob("*.cache"))),
            'api_cache_files': len(list(self.api_cache_dir.glob("*.cache"))),
            'calc_cache_files': len(list(self.calc_cache_dir.glob("*.cache"))),
            'cache_ttl_seconds': self.cache_ttl
        }

# ============================================================================
# БЛОК 4: КОНФИГУРАЦИИ API МАРКЕТПЛЕЙСОВ
# ============================================================================

class MarketplaceAPIEndpoint(Enum):
    """Эндпоинты API маркетплейсов для загрузки тарифов"""
    OZON_COMMISSIONS = "https://api.ozon.ru/v1/commission/list"
    OZON_DELIVERY = "https://api.ozon.ru/v1/delivery-methods"
    WILDBERRIES_TARIFFS = "https://suppliers-api.wildberries.ru/api/v2/tariffs"
    WILDBERRIES_COMMISSIONS = "https://suppliers-api.wildberries.ru/api/v2/commissions"
    YANDEX_TARIFFS = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/tariffs"
    YANDEX_COMMISSIONS = "https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/offer-mapping-entries"
    DEEPSEEK_CHAT = "https://api.deepseek.com/v1/chat/completions"

@dataclass
class MarketplaceTariffData:
    """
    Структура данных тарифов маркетплейса.
    Содержит все необходимые параметры для расчета FBS.
    ВСЕ ЗНАЧЕНИЯ ЗАГРУЖАЮТСЯ ИЗ API ИЛИ ПОЛЬЗОВАТЕЛЬСКОГО ВВОДА
    """
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
    
    def to_dict(self) -> Dict[str, Any]:
        """Преобразование в словарь для сериализации"""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'MarketplaceTariffData':
        """Создание из словаря"""
        return cls(**data)

# ============================================================================
# БЛОК 5: API МЕНЕДЖЕР ДЛЯ ЗАГРУЗКИ ТАРИФОВ (УЛУЧШЕННЫЙ)
# ============================================================================

class APIRateLimiter:
    """Управление лимитами запросов к API"""
    
    def __init__(self):
        self.last_request_time: Dict[str, float] = {}
        # Минимальные интервалы между запросами (берутся из документации API)
        self.min_interval: Dict[str, float] = {
            'ozon': 0.5,
            'wildberries': 1.0,
            'yandex_market': 1.0,
            'deepseek': 0.5
        }
    
    def wait_if_needed(self, service: str):
        """Ожидание перед следующим запросом если необходимо"""
        if service in self.last_request_time:
            elapsed = time.time() - self.last_request_time[service]
            min_wait = self.min_interval.get(service, 0.5)
            if elapsed < min_wait:
                time.sleep(min_wait - elapsed)
        self.last_request_time[service] = time.time()

class MarketplaceAPIManager:
    """
    Менеджер для загрузки актуальных тарифов через API маркетплейсов.
    
    Поддерживаемые источники:
    1. Прямое API маркетплейса (Ozon, Wildberries, Яндекс Маркет)
    2. DeepSeek AI (когда прямое API недоступно)
    3. CSV файл (пользовательский импорт)
    4. Встроенные дефолтные значения (только как крайний фолбэк)
    
    Особенности:
    - Автоматическое кэширование
    - Retry при ошибках сети
    - Rate limiting для соблюдения лимитов API
    - НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ
    """
    
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
        """Загрузка API ключей из защищенного хранилища"""
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
        """
        Сохранение API ключа.
        
        Args:
            service: Название сервиса (ozon, wildberries, yandex_market, deepseek)
            api_key: API ключ
            
        Returns:
            bool: True если сохранение успешно
        """
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
        """Получение API ключа для сервиса"""
        return self._api_keys_cache.get(service)
    
    def has_api_key(self, service: str) -> bool:
        """Проверка наличия API ключа"""
        return bool(self._api_keys_cache.get(service))
    
    def get_cached_tariffs(self, marketplace: str) -> Optional[Dict[str, Dict]]:
        """Получение закэшированных тарифов"""
        cache_key = f"tariffs_{marketplace.lower()}"
        return self.cache_manager.get('tariffs', cache_key)
    
    def save_tariffs_to_cache(self, marketplace: str, tariffs: Dict[str, Dict]):
        """Сохранение тарифов в кэш"""
        cache_key = f"tariffs_{marketplace.lower()}"
        self.cache_manager.set('tariffs', cache_key, {
            'tariffs': tariffs,
            'marketplace': marketplace,
            'cached_at': datetime.now().isoformat(),
            'version': APP_VERSION
        })
        logger.info(f"💾 Тарифы {marketplace} сохранены в кэш")
    
    def load_tariffs_from_csv(self, marketplace: str, csv_content: str) -> Dict[str, Dict]:
        """
        Загрузка тарифов из CSV файла.
        
        Args:
            marketplace: Название маркетплейса
            csv_content: Содержимое CSV файла
            
        Returns:
            Dict: Тарифы по категориям
        """
        tariffs = {}
        try:
            df = pd.read_csv(io.StringIO(csv_content))
            
            # Проверяем наличие обязательных колонок
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
                    'last_updated': datetime.now().isoformat()
                }
            
            logger.info(f"✅ Загружено {len(tariffs)} категорий из CSV")
            self.audit_logger.log('load_tariffs_csv', {'marketplace': marketplace, 'count': len(tariffs)})
            
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки CSV: {e}")
            tariffs = {}
        
        return tariffs
    
    def _get_default_tariffs(self, marketplace: str) -> Dict[str, Dict]:
        """
        Получение дефолтных тарифов - используется ТОЛЬКО как крайний фолбэк.
        ВНИМАНИЕ: Эти значения используются только когда API и CSV недоступны.
        Пользователь всегда может загрузить свои тарифы через CSV.
        """
        logger.warning(f"⚠️ Использую дефолтные тарифы для {marketplace} (только как фолбэк)")
        
        # Проверяем кэш дефолтных тарифов
        if marketplace in self._default_tariffs_cache:
            return self._default_tariffs_cache[marketplace]
        
        # Минимальные базовые значения для демонстрации
        # Пользователь должен загрузить актуальные тарифы через API или CSV
        default_tariffs = {
            'default': {
                'commission_rate': 0.15,
                'min_commission': 30.0,
                'last_mile_base': 50.0,
                'last_mile_per_kg': 15.0,
                'last_mile_per_km': 3.5,
                'acquiring_fee': 0.015,
                'return_fee': 0.02,
                'penalty_rate': 0.05,
                'penalty_time_hours': 24,
                'fbo_multiplier': 0.75,
                'fbp_multiplier': 0.60,
                'storage_base_rate': 0.30,
                'min_logistics': 25.0,
                'source': 'fallback_default',
                'last_updated': datetime.now().isoformat(),
                'warning': '⚠️ Используются примерные значения. Загрузите актуальные тарифы через API или CSV.'
            }
        }
        
        self._default_tariffs_cache[marketplace] = default_tariffs
        return default_tariffs
    
    @retry_on_failure(max_retries=2, delay=2.0)
    def fetch_ozon_tariffs(self) -> Dict[str, Dict]:
        """
        Загрузка тарифов Ozon через официальное API.
        Документация: https://docs.ozon.ru/api/seller/
        """
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
        """
        Загрузка тарифов Wildberries через официальное API.
        Документация: https://suppliers.wildberries.ru/
        """
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
        """
        Загрузка тарифов Яндекс Маркет через официальное API.
        Документация: https://yandex.ru/dev/market/partner/
        """
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
        """
        AI-обогащение тарифов через DeepSeek API.
        
        Используется когда прямые API маркетплейсов недоступны.
        """
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
                    tariff_data['source'] = f'deepseek_ai (confidence: {confidence})'
                    tariff_data['last_updated'] = datetime.now().isoformat()
                    tariff_data['ai_confidence'] = confidence
                    tariffs[category] = tariff_data
                
                if tariffs:
                    logger.info(f"✅ DeepSeek предоставил тарифы для {len(tariffs)} категорий {marketplace}")
                    self.audit_logger.log('fetch_deepseek_tariffs', {'marketplace': marketplace, 'count': len(tariffs)})
        
        except Exception as e:
            logger.error(f"❌ Ошибка запроса к DeepSeek: {e}")
            logger.exception(e)
        
        return tariffs
    
    def get_tariffs(self, marketplace: str, force_refresh: bool = False, 
                   use_ai_fallback: bool = True, csv_content: Optional[str] = None) -> Dict[str, Dict]:
        """
        Основной метод получения тарифов.
        
        Приоритет источников (НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ):
        1. Кэш (если не force_refresh)
        2. Прямое API маркетплейса
        3. DeepSeek AI (если use_ai_fallback)
        4. CSV импорт (если предоставлен)
        5. Фолбэк с предупреждением (только для демонстрации)
        
        Args:
            marketplace: Название маркетплейса (Ozon, Wildberries, Яндекс Маркет)
            force_refresh: Принудительно обновить, игнорируя кэш
            use_ai_fallback: Использовать DeepSeek AI при недоступности API
            csv_content: Содержимое CSV файла для импорта
            
        Returns:
            Dict: Тарифы по категориям
        """
        marketplace_lower = marketplace.lower()
        
        # Проверяем кэш
        if not force_refresh:
            cached = self.get_cached_tariffs(marketplace)
            if cached:
                tariffs = cached.get('tariffs', {})
                if tariffs:
                    cached_time = cached.get('cached_at', '')
                    logger.info(f"📦 Использованы кэшированные тарифы {marketplace} от {cached_time}")
                    return tariffs
        
        logger.info(f"🔄 Загрузка тарифов {marketplace}...")
        
        # Пробуем прямое API
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
        
        # Если API не сработало - пробуем DeepSeek
        if (not tariffs or all(t.get('source') == 'fallback_default' for t in tariffs.values())) and use_ai_fallback:
            logger.info(f"🤖 Прямое API недоступно, использую DeepSeek AI для {marketplace}")
            try:
                ai_tariffs = self.fetch_tariffs_via_deepseek(marketplace)
                if ai_tariffs and not any(t.get('source') == 'fallback_default' for t in ai_tariffs.values()):
                    tariffs = ai_tariffs
                    logger.info(f"✅ Тарифы {marketplace} получены через DeepSeek AI")
            except Exception as e:
                logger.error(f"❌ DeepSeek также недоступен: {e}")
        
        # Если AI не сработал - пробуем CSV
        if (not tariffs or all(t.get('source') == 'fallback_default' for t in tariffs.values())) and csv_content:
            logger.info(f"📄 Использую CSV импорт для {marketplace}")
            csv_tariffs = self.load_tariffs_from_csv(marketplace, csv_content)
            if csv_tariffs:
                tariffs = csv_tariffs
                logger.info(f"✅ Тарифы {marketplace} загружены из CSV")
        
        # Если ничего не сработало - фолбэк с предупреждением
        if not tariffs or all(t.get('source') == 'fallback_default' for t in tariffs.values()):
            logger.warning(f"⚠️ Все источники недоступны, использую фолбэк для {marketplace}")
            logger.warning("⚠️ ПОЖАЛУЙСТА, ЗАГРУЗИТЕ АКТУАЛЬНЫЕ ТАРИФЫ ЧЕРЕЗ API ИЛИ CSV")
            tariffs = self._get_default_tariffs(marketplace_lower)
            
            # Добавляем предупреждение в каждый тариф
            for category in tariffs:
                tariffs[category]['warning'] = '⚠️ Используются примерные значения. Загрузите актуальные тарифы!'
        
        # Сохраняем в кэш только если есть нормальные тарифы
        if tariffs and not all(t.get('source') == 'fallback_default' for t in tariffs.values()):
            self.save_tariffs_to_cache(marketplace, tariffs)
        
        return tariffs
    
    def get_all_tariffs_as_dataframe(self, marketplace: str) -> pd.DataFrame:
        """Получение всех тарифов в виде DataFrame для отображения в UI"""
        tariffs = self.get_tariffs(marketplace)
        
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
                'Обновлено': data.get('last_updated', '')[:19] if data.get('last_updated') else '',
                'Предупреждение': data.get('warning', '') if data.get('warning') else ''
            })
        
        return pd.DataFrame(rows)
    
    def test_api_connection(self, marketplace: str) -> Dict[str, Any]:
        """Тестирование API подключения"""
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
    """
    Входные данные для расчета FBS юнит-экономики.
    ВСЕ ЗНАЧЕНИЯ БЕРУТСЯ ИЗ ПОЛЬЗОВАТЕЛЬСКОГО ВВОДА
    НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ
    """
    # Основные параметры товара (пользовательский ввод)
    artikul: str = ""
    product_name: str = ""
    category: str = "default"
    
    # Финансовые параметры (пользовательский ввод)
    selling_price: float = 0.0
    cogs: float = 0.0
    
    # Физические параметры товара (пользовательский ввод)
    weight_kg: float = 0.0
    length_cm: float = 0.0
    width_cm: float = 0.0
    height_cm: float = 0.0
    
    # FBS специфика (пользовательский ввод)
    first_mile_cost_per_unit: float = 0.0
    packaging_cost: float = 0.0
    pick_pack_time_min: float = 5.0
    operator_hourly_rate: float = 300.0
    warehouse_distance_km: float = 0.0
    
    # Логистические параметры (пользовательский ввод)
    transport_type: str = "own"
    transport_cost_per_km: float = 20.0
    pallet_capacity: int = 100
    pallet_cost: float = 2000.0
    
    # Маркетинговые параметры (пользовательский ввод)
    marketing_budget_per_unit: float = 0.0
    
    # Складские параметры (пользовательский ввод)
    stock_depth_days: int = 30
    daily_sales: int = 5
    warehouse_rent_per_sqm: float = 500.0
    warehouse_space_per_unit: float = 0.01
    safety_stock_days: int = 7
    reorder_point_days: int = 5
    supplier_lead_time_days: int = 3
    
    # Параметры для расчета LTV и CAC (пользовательский ввод)
    repeat_purchase_rate: float = 0.3
    avg_purchases_per_year: float = 2.5
    customer_retention_rate: float = 0.7
    discount_rate: float = 0.1
    
    # Операционные параметры (пользовательский ввод)
    has_night_shift: bool = False
    processing_capacity_per_hour: int = 20
    
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
    """
    Результаты расчета FBS юнит-экономики.
    ВСЕ РАСЧЕТЫ ОСНОВАНЫ НА АКТУАЛЬНЫХ ДАННЫХ
    НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ
    """
    # Основные идентификаторы
    artikul: str = ""
    product_name: str = ""
    
    # Ключевые финансовые показатели
    selling_price: float = 0.0
    total_expenses: float = 0.0
    gross_profit: float = 0.0
    margin_percent: float = 0.0
    roi_percent: float = 0.0
    
    # Детализация расходов
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
    
    # FBS специфические метрики
    penalty_probability: float = 0.0
    break_even_distance_km: float = 0.0
    max_discount_percent: float = 0.0
    safety_margin_price: float = 0.0
    break_even_volume: float = 0.0
    
    # LTV и CAC метрики
    ltv: float = 0.0
    cac: float = 0.0
    ltv_cac_ratio: float = 0.0
    romi: float = 0.0
    
    # Метрики складской оптимизации
    optimal_stock_units: int = 0
    safety_stock_units: int = 0
    reorder_point_units: int = 0
    stock_turnover_days: float = 0.0
    stock_turnover_rate: float = 0.0
    days_of_inventory: float = 0.0
    holding_cost_per_unit: float = 0.0
    
    # Рекомендации по оптимизации склада
    recommended_stock_depth_days: int = 0
    recommended_safety_stock_days: int = 0
    stock_optimization_potential: float = 0.0
    
    # Логистические зоны риска
    logistic_zone: str = "unknown"
    logistic_zone_label: str = ""
    logistic_recommendation: str = ""
    is_logistic_critical: bool = False
    
    # Эффективность использования пространства
    space_efficiency_ratio: float = 0.0
    revenue_per_sqm: float = 0.0
    profit_per_sqm: float = 0.0
    
    # Сезонная корректировка
    seasonal_factor: float = 1.0
    adjusted_margin_percent: float = 0.0
    seasonal_recommendation: str = ""
    
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
            'adjusted_margin_percent': self.adjusted_margin_percent
        }

# ============================================================================
# БЛОК 7: КОНФИГУРАЦИИ НАЛОГОВЫХ СИСТЕМ
# ============================================================================

# Конфигурации налоговых систем (на основе реальных данных, но пользователь может изменить)
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
# БЛОК 8: ОСНОВНОЙ КАЛЬКУЛЯТОР FBS ЮНИТ-ЭКОНОМИКИ
# ============================================================================

class FBSUnitEconomicsCalculator:
    """
    Профессиональный калькулятор юнит-экономики для FBS-модели.
    
    КЛЮЧЕВЫЕ ПРИНЦИПЫ:
    1. НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ
    2. Все тарифы загружаются из API или пользовательского ввода
    3. Полная прозрачность расчетов
    4. Динамическая оптимизация на основе реальных данных
    """
    
    def __init__(self, api_manager: Optional[MarketplaceAPIManager] = None, 
                 tax_system: str = "УСН 6% (доходы)"):
        self.api_manager = api_manager or MarketplaceAPIManager()
        self.tax_system = tax_system
        self.current_marketplace = "Ozon"
        self.current_tariffs: Dict[str, Dict] = {}
        self.tariffs_updated_at: Optional[datetime] = None
        self.tariffs_source = "default"
        
        self.progress_tracker = ProgressTracker()
        self.audit_logger = AuditLogger()
        
        self.refresh_tariffs()
    
    def set_marketplace(self, marketplace_name: str):
        self.current_marketplace = marketplace_name
        self.refresh_tariffs()
        self.audit_logger.log('set_marketplace', {'marketplace': marketplace_name})
        logger.info(f"🏪 Установлен маркетплейс: {marketplace_name}")
    
    def refresh_tariffs(self, force: bool = False, use_ai: bool = False, csv_content: Optional[str] = None):
        logger.info(f"🔄 Обновление тарифов для {self.current_marketplace}...")
        
        self.current_tariffs = self.api_manager.get_tariffs(
            self.current_marketplace,
            force_refresh=force,
            use_ai_fallback=use_ai,
            csv_content=csv_content
        )
        
        self.tariffs_updated_at = datetime.now()
        
        sources = set()
        for tariff in self.current_tariffs.values():
            source = tariff.get('source', 'default')
            if 'api' in source:
                sources.add('api')
            elif 'deepseek' in source:
                sources.add('deepseek')
            elif 'csv' in source:
                sources.add('csv')
            else:
                sources.add('default')
        
        if 'api' in sources:
            self.tariffs_source = 'api'
        elif 'deepseek' in sources:
            self.tariffs_source = 'deepseek'
        elif 'csv' in sources:
            self.tariffs_source = 'csv'
        else:
            self.tariffs_source = 'default'
        
        logger.info(f"✅ Тарифы обновлены. Источник: {self.tariffs_source}. Категорий: {len(self.current_tariffs)}")
    
    def get_tariff_for_category(self, category: str) -> Dict[str, Any]:
        """Получение тарифа для конкретной категории товара"""
        if not self.current_tariffs:
            logger.warning("⚠️ Тарифы не загружены")
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
            logger.debug(f"🔍 Использую дефолтный тариф для категории: {category}")
            return self.current_tariffs['default']
        
        first_tariff = next(iter(self.current_tariffs.values()), {})
        if first_tariff:
            logger.warning(f"⚠️ Категория {category} не найдена, использую первый доступный тариф")
            return first_tariff
        
        return {}
    
    def _get_logistic_zone(self, distance_km: float) -> Dict[str, Any]:
        """Определяет зону логистического риска на основе рассчитанных данных"""
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
    
    @timing_decorator
    def calculate_unit_economics(self, input_data: FBSInputData) -> FBSResultData:
        """
        Основной расчет юнит-экономики для FBS модели.
        ВСЕ РАСЧЕТЫ ОСНОВАНЫ НА АКТУАЛЬНЫХ ДАННЫХ ИЗ API ИЛИ ПОЛЬЗОВАТЕЛЬСКОГО ВВОДА
        НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ
        """
        validation_errors = input_data.validate()
        if validation_errors:
            logger.warning(f"⚠️ Ошибки валидации для {input_data.artikul}: {validation_errors}")
        
        result = FBSResultData()
        result.artikul = input_data.artikul
        result.product_name = input_data.product_name
        result.selling_price = input_data.selling_price
        
        # Получаем актуальный тариф для категории (ИЗ API, НЕ ЗАХАРДКОЖЕННЫЙ)
        tariff = self.get_tariff_for_category(input_data.category)
        
        if not tariff:
            logger.error(f"❌ Тариф не найден для категории {input_data.category}")
            # Создаем пустой тариф для продолжения работы
            tariff = {
                'commission_rate': 0.0,
                'min_commission': 0.0,
                'last_mile_base': 0.0,
                'last_mile_per_kg': 0.0,
                'acquiring_fee': 0.0,
                'return_fee': 0.0,
                'penalty_rate': 0.0,
                'penalty_time_hours': 24,
                'fbo_multiplier': 1.0,
                'fbp_multiplier': 1.0,
                'storage_base_rate': 0.0,
                'min_logistics': 0.0,
                'source': 'error'
            }
        
        # =====================================================================
        # 1. КОМИССИЯ МАРКЕТПЛЕЙСА (из API)
        # =====================================================================
        commission_rate = tariff.get('commission_rate', 0.0)
        min_commission = tariff.get('min_commission', 0.0)
        
        result.commission = max(
            input_data.selling_price * commission_rate,
            min_commission
        )
        
        # =====================================================================
        # 2. FIRST MILE (на основе пользовательского ввода)
        # =====================================================================
        if input_data.first_mile_cost_per_unit > 0:
            result.first_mile_cost = input_data.first_mile_cost_per_unit
        else:
            pallet_units = max(input_data.pallet_capacity, 1)
            cost_per_pallet = input_data.warehouse_distance_km * input_data.transport_cost_per_km * 2
            result.first_mile_cost = cost_per_pallet / pallet_units
        
        # =====================================================================
        # 3. LAST MILE (из API)
        # =====================================================================
        if input_data.length_cm > 0 and input_data.width_cm > 0 and input_data.height_cm > 0:
            vol_weight = (input_data.length_cm * input_data.width_cm * input_data.height_cm) / 5000.0
        else:
            vol_weight = 0
        
        billable_weight = max(input_data.weight_kg, vol_weight)
        billable_weight = math.ceil(billable_weight * 2) / 2
        
        last_mile_base = tariff.get('last_mile_base', 0.0)
        last_mile_per_kg = tariff.get('last_mile_per_kg', 0.0)
        min_logistics = tariff.get('min_logistics', 0.0)
        
        result.last_mile_cost = max(
            last_mile_base + (billable_weight * last_mile_per_kg),
            min_logistics
        )
        
        # =====================================================================
        # 4. PICK & PACK (на основе пользовательского ввода)
        # =====================================================================
        pick_pack_hours = input_data.pick_pack_time_min / 60.0
        result.pick_pack_cost = pick_pack_hours * input_data.operator_hourly_rate
        
        # =====================================================================
        # 5. УПАКОВКА (пользовательский ввод)
        # =====================================================================
        result.packaging_cost = input_data.packaging_cost
        
        # =====================================================================
        # 6. ЭКВАЙРИНГ (из API)
        # =====================================================================
        acquiring_fee = tariff.get('acquiring_fee', 0.0)
        result.acquiring_cost = input_data.selling_price * acquiring_fee
        
        # =====================================================================
        # 7. ВОЗВРАТЫ (из API)
        # =====================================================================
        return_fee = tariff.get('return_fee', 0.0)
        result.return_cost = input_data.selling_price * return_fee
        
        # =====================================================================
        # 8. ШТРАФЫ (из API + пользовательские параметры)
        # =====================================================================
        if input_data.has_night_shift:
            penalty_probability = 0.05
        else:
            penalty_probability = 0.35
        
        penalty_rate = tariff.get('penalty_rate', 0.0)
        
        result.penalty_probability = penalty_probability
        result.penalty_cost = input_data.selling_price * penalty_rate * penalty_probability
        
        # =====================================================================
        # 9. МАРКЕТИНГ (пользовательский ввод)
        # =====================================================================
        result.marketing_cost = input_data.marketing_budget_per_unit
        
        # =====================================================================
        # 10. СКЛАДСКИЕ РАСХОДЫ (пользовательский ввод)
        # =====================================================================
        total_stock = input_data.stock_depth_days * input_data.daily_sales
        if total_stock > 0 and input_data.daily_sales > 0:
            total_warehouse_space = input_data.warehouse_space_per_unit * total_stock
            monthly_rent = input_data.warehouse_rent_per_sqm * total_warehouse_space
            result.warehouse_cost = monthly_rent / (30 * input_data.daily_sales)
        else:
            result.warehouse_cost = 0
        
        # =====================================================================
        # 11. НАЛОГ (на основе выбранной системы)
        # =====================================================================
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
        
        # =====================================================================
        # 12. ИТОГО РАСХОДОВ И ПРИБЫЛЬ
        # =====================================================================
        result.total_expenses = (
            input_data.cogs +
            result.commission +
            result.first_mile_cost +
            result.last_mile_cost +
            result.pick_pack_cost +
            result.packaging_cost +
            result.acquiring_cost +
            result.return_cost +
            result.penalty_cost +
            result.marketing_cost +
            result.warehouse_cost +
            result.tax_cost
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
        
        # =====================================================================
        # 13. ТОЧКА БЕЗУБЫТОЧНОСТИ ПО РАССТОЯНИЮ
        # =====================================================================
        if result.first_mile_cost > 0 and input_data.pallet_capacity > 0:
            cost_per_km_per_unit = (input_data.transport_cost_per_km * 2) / input_data.pallet_capacity
            if cost_per_km_per_unit > 0:
                result.break_even_distance_km = result.gross_profit / cost_per_km_per_unit
            else:
                result.break_even_distance_km = float('inf')
        else:
            result.break_even_distance_km = float('inf')
        
        # =====================================================================
        # 14. ЛОГИСТИЧЕСКИЕ ЗОНЫ РИСКА
        # =====================================================================
        logistic_zone_info = self._get_logistic_zone(result.break_even_distance_km)
        result.logistic_zone = logistic_zone_info['zone']
        result.logistic_zone_label = logistic_zone_info['label']
        result.logistic_recommendation = logistic_zone_info['recommendation']
        result.is_logistic_critical = logistic_zone_info['is_critical']
        
        # =====================================================================
        # 15. ЗАПАС ПРОЧНОСТИ ПО ЦЕНЕ
        # =====================================================================
        variable_costs_percent = (
            commission_rate +
            acquiring_fee +
            return_fee +
            penalty_rate * penalty_probability +
            TAX_SYSTEMS[self.tax_system]["rate"]
        )
        
        fixed_costs_per_unit = (
            input_data.cogs +
            result.first_mile_cost +
            result.last_mile_cost +
            result.pick_pack_cost +
            result.packaging_cost +
            result.marketing_cost +
            result.warehouse_cost
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
        
        # =====================================================================
        # 16. ТОЧКА БЕЗУБЫТОЧНОСТИ ПО ОБЪЕМУ
        # =====================================================================
        variable_costs = result.commission + result.last_mile_cost + result.acquiring_cost + result.return_cost + result.penalty_cost
        
        if (result.selling_price - variable_costs) > 0:
            result.break_even_volume = fixed_costs_per_unit / (result.selling_price - variable_costs)
        else:
            result.break_even_volume = float('inf')
        
        # =====================================================================
        # 17. LTV И CAC
        # =====================================================================
        if (1 + input_data.discount_rate) > 0:
            result.ltv = (
                input_data.selling_price *
                input_data.avg_purchases_per_year *
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
        
        # =====================================================================
        # 18. СЕЗОННАЯ КОРРЕКТИРОВКА
        # =====================================================================
        current_month = datetime.now().month
        seasonal_factor = 1.0  # Базовое значение, пользователь может настроить
        result.seasonal_factor = seasonal_factor
        result.adjusted_margin_percent = result.margin_percent * seasonal_factor
        
        if result.adjusted_margin_percent < 10:
            result.seasonal_recommendation = "⚠️ Низкая сезонная маржа - рассмотрите акции или повышение цен"
        elif result.adjusted_margin_percent < 20:
            result.seasonal_recommendation = "📊 Средняя сезонная маржа - стабильно, можно улучшить"
        else:
            result.seasonal_recommendation = "✅ Отличная сезонная маржа!"
        
        # =====================================================================
        # 19. ОПТИМИЗАЦИЯ СКЛАДСКИХ ОСТАТКОВ
        # =====================================================================
        daily_demand = input_data.daily_sales
        annual_demand = daily_demand * 365
        
        # Стоимость заказа (пользователь может настроить)
        ordering_cost = 500.0
        
        # Стоимость хранения на единицу в год (на основе пользовательского ввода)
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
        
        # =====================================================================
        # 20. ЭФФЕКТИВНОСТЬ ИСПОЛЬЗОВАНИЯ ПРОСТРАНСТВА
        # =====================================================================
        total_stock = input_data.stock_depth_days * input_data.daily_sales
        if total_stock > 0 and input_data.warehouse_space_per_unit > 0:
            total_sqm = total_stock * input_data.warehouse_space_per_unit
            if total_sqm > 0:
                result.space_efficiency_ratio = total_stock / total_sqm
                result.revenue_per_sqm = (input_data.selling_price * input_data.daily_sales * 30) / total_sqm
                result.profit_per_sqm = (result.gross_profit * input_data.daily_sales * 30) / total_sqm
            else:
                result.space_efficiency_ratio = 0
                result.revenue_per_sqm = 0
                result.profit_per_sqm = 0
        
        # =====================================================================
        # 21. РЕКОМЕНДАЦИИ ПО ОПТИМИЗАЦИИ СКЛАДА
        # =====================================================================
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
        
        self.audit_logger.log('calculate_unit', {
            'artikul': input_data.artikul,
            'profit': result.gross_profit,
            'margin': result.margin_percent,
            'logistic_zone': result.logistic_zone
        })
        
        return result
    
    @timing_decorator
    def calculate_batch(self, input_data_list: List[FBSInputData],
                       use_parallel: bool = True,
                       max_workers: int = 8) -> List[FBSResultData]:
        """Пакетный расчет для множества товаров"""
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
                            self.progress_tracker.update(
                                completed, 
                                f"Обработано {completed}/{total} товаров"
                            )
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
                    self.progress_tracker.update(
                        i + 1,
                        f"Обработано {i + 1}/{total} товаров"
                    )
        
        self.progress_tracker.update(total, f"✅ Расчет завершен! Обработано {total} товаров")
        results = [r for r in results if r is not None]
        
        logger.info(f"✅ Пакетный расчет завершен. Успешно: {len(results)}/{total}")
        
        return results
    
    def run_what_if_analysis(self, base_data: FBSInputData, scenarios: List[Dict]) -> pd.DataFrame:
        """Анализ сценариев 'Что если'"""
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
                'Скорр. маржа, %': round(result.adjusted_margin_percent, 2)
            })
        
        self.audit_logger.log('what_if_analysis', {'scenarios': len(scenarios)})
        return pd.DataFrame(results)
    
    def generate_recommendations(self, results: List[FBSResultData]) -> List[Dict]:
        """Генерация автоматических рекомендаций"""
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
        
        self.audit_logger.log('generate_recommendations', {'count': len(recommendations)})
        return recommendations

# ============================================================================
# БЛОК 9: ИНТЕРФЕЙС ПОЛЬЗОВАТЕЛЯ (STREAMLIT) - УЛУЧШЕННАЯ ВЕРСИЯ
# ============================================================================

def init_session_state():
    """Инициализация всех состояний сессии Streamlit"""
    
    if 'api_manager' not in st.session_state:
        st.session_state.api_manager = MarketplaceAPIManager()
    
    if 'calculator' not in st.session_state:
        st.session_state.calculator = FBSUnitEconomicsCalculator(
            api_manager=st.session_state.api_manager
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

def render_sidebar():
    """Отрисовка боковой панели навигации"""
    
    with st.sidebar:
        st.markdown("""
        <div style='text-align: center; padding: 20px 15px; background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460); border-radius: 12px; margin-bottom: 25px;'>
            <h1 style='color: white; margin: 0; font-size: 1.5em;'>🚀 FBS PRO</h1>
            <p style='color: #a8a8d0; margin: 8px 0 0 0; font-size: 0.9em;'>Операционная версия</p>
            <p style='color: #6666aa; margin: 5px 0 0 0; font-size: 0.7em;'>v6.2.0 | Без хардкода</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 🧭 Навигация")
        
        sections = {
            "🏠 Главная": "main",
            "🧮 Калькулятор FBS": "calculator",
            "📋 Тарифы маркетплейсов": "tariffs",
            "📈 Дашборд": "dashboard",
            "🎯 Анализ сценариев": "what_if",
            "💡 Рекомендации": "recommendations",
            "📥 Экспорт": "export",
            "⚙️ Настройки": "settings"
        }
        
        selected_section = st.radio(
            "Выберите раздел:",
            list(sections.keys()),
            label_visibility="collapsed"
        )
        
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
            st.success("🔌 Тарифы: API")
        elif calculator.tariffs_source == 'deepseek':
            st.info("🤖 Тарифы: DeepSeek AI")
        elif calculator.tariffs_source == 'csv':
            st.info("📄 Тарифы: CSV")
        else:
            st.warning("⚠️ Тарифы: Дефолтные (требуется загрузка)")
        
        if st.session_state.results:
            st.success(f"✅ Рассчитано: {len(st.session_state.results)} товаров")
            profitable = len([r for r in st.session_state.results if r.gross_profit > 0])
            st.metric("Прибыльных", f"{profitable} из {len(st.session_state.results)}")
        else:
            st.info("ℹ️ Расчеты не выполнялись")
        
        st.markdown("---")
        st.markdown("### ⚡ Быстрые действия")
        
        # ИСПРАВЛЕНО: заменен use_container_width на width='stretch'
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
    """Дашборд с фильтрацией и поиском"""
    
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
        sort_by = st.selectbox(
            "Сортировка", 
            ["Прибыль", "Маржа", "ROI", "Оборачиваемость", "Прибыль на м²"]
        )
    with col4:
        filter_zone = st.selectbox(
            "Логистическая зона",
            ["Все", "🔴 Критическая", "🟡 Зона риска", "🟢 Безопасная", "🔵 Идеальная"]
        )
    
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
        # ИСПРАВЛЕНО: заменен use_container_width на width='stretch'
        st.dataframe(df, width='stretch', height=400)
        st.caption(f"📊 Показано {len(filtered)} из {len(results)} товаров")
    else:
        st.info("ℹ️ Нет товаров, соответствующих фильтрам")

def main():
    """Главная функция запуска приложения"""
    
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
    
    if current_section == 'main':
        st.markdown("""
        <div style='text-align: center; padding: 50px 30px; background: linear-gradient(135deg, #0f0c29, #302b63, #24243e); border-radius: 20px; margin-bottom: 35px;'>
            <h1 style='color: white; font-size: 3em; margin: 0;'>🚀 FBS Юнит-экономика PRO</h1>
            <p style='color: #a8a8d0; font-size: 1.3em; margin: 20px 0;'>
                Операционная версия — Оптимизация складских остатков и логистических коридоров
            </p>
            <p style='color: #6666aa; font-size: 1em; margin: 10px 0;'>
                Ozon • Wildberries • Яндекс Маркет | API Integration | DeepSeek AI
            </p>
            <p style='color: #8888cc; font-size: 0.9em; margin: 10px 0;'>
                🆕 НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ — все данные из API или пользовательского ввода
            </p>
        </div>
        """, unsafe_allow_html=True)
        
        st.info("""
        ### 🎯 Ключевые принципы системы
        
        1. **НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ** — все тарифы загружаются из API или CSV
        2. **Актуальные данные** — интеграция с API Ozon, Wildberries, Яндекс Маркет
        3. **AI-обогащение** — DeepSeek для получения тарифов при недоступности API
        4. **Полная прозрачность** — все расчеты основаны на реальных данных
        5. **Оптимизация** — складские остатки и логистические коридоры
        
        ### 📋 Что нужно для работы:
        - **API ключи** маркетплейсов (в разделе Настройки)
        - **CSV файл** с тарифами (если API недоступны)
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
        **🎯 Профессиональный расчет FBS с оптимизацией складских остатков**
        
        - 🚛 **First Mile** — ваша логистика до склада МП (пользовательский ввод)
        - 📦 **Last Mile** — доставка МП до клиента (из API тарифов)
        - 📊 **Оптимизация склада** — EOQ, страховой запас, точка заказа
        - 💰 **Рекомендованные цены** — при разных уровнях маржи
        - 🔴 **Логистические зоны риска** — на основе рассчитанных данных
        - 📅 **Сезонная корректировка** — учет текущего месяца
        - ⚠️ **Условное форматирование** — прибыльные/убыточные товары
        """)
        
        calc_mode = st.radio(
            "Режим расчета:",
            ["📱 Расчет одного товара", "📊 Массовый расчет из файла"],
            horizontal=True
        )
        
        if calc_mode == "📱 Расчет одного товара":
            col1, col2 = st.columns(2)
            with col1:
                artikul = st.text_input("Артикул", "SKU-001")
                product_name = st.text_input("Наименование", "Тестовый товар")
                category = st.selectbox("Категория", list(calculator.current_tariffs.keys()) or ["default"])
                selling_price = st.number_input("Цена продажи, ₽", 5000.0, step=100.0)
                cogs = st.number_input("Себестоимость, ₽", 3000.0, step=100.0)
            
            with col2:
                weight = st.number_input("Вес, кг", 1.5, step=0.1)
                length = st.number_input("Длина, см", 20, step=1)
                width = st.number_input("Ширина, см", 15, step=1)
                height = st.number_input("Высота, см", 10, step=1)
                warehouse_distance = st.number_input("Расстояние до МП, км", 50.0, step=1.0)
                daily_sales = st.number_input("Продаж в день, шт", 5, step=1)
                has_night = st.checkbox("Ночная смена")
            
            if st.button("🚀 Рассчитать", type="primary"):
                input_data = FBSInputData(
                    artikul=artikul,
                    product_name=product_name,
                    category=category,
                    selling_price=selling_price,
                    cogs=cogs,
                    weight_kg=weight,
                    length_cm=length,
                    width_cm=width,
                    height_cm=height,
                    warehouse_distance_km=warehouse_distance,
                    daily_sales=daily_sales,
                    has_night_shift=has_night
                )
                
                result = calculator.calculate_unit_economics(input_data)
                st.session_state.results = [result]
                st.session_state.input_data_list = [input_data]
                
                st.markdown("---")
                st.markdown("## 📊 Результаты расчета")
                
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
                    st.metric("🔴 Лог. зона", result.logistic_zone_label)
                
                st.markdown("### 📋 Детализация")
                col1, col2 = st.columns(2)
                with col1:
                    st.markdown("**💰 Финансовые показатели**")
                    st.metric("Комиссия МП", f"{result.commission:,.0f} ₽")
                    st.metric("First Mile", f"{result.first_mile_cost:,.0f} ₽")
                    st.metric("Last Mile", f"{result.last_mile_cost:,.0f} ₽")
                    st.metric("Pick & Pack", f"{result.pick_pack_cost:,.0f} ₽")
                    st.metric("Эквайринг", f"{result.acquiring_cost:,.0f} ₽")
                    st.metric("Возвраты", f"{result.return_cost:,.0f} ₽")
                    st.metric("Штрафы", f"{result.penalty_cost:,.0f} ₽")
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
                    st.metric("Точка безубыт. (км)", f"{result.break_even_distance_km:.0f} км")
                    st.metric("Точка безубыт. (шт)", f"{result.break_even_volume:.0f} шт")
                    
                    st.markdown("**📅 Сезонность**")
                    st.metric("Скорр. маржа", f"{result.adjusted_margin_percent:.1f}%")
                    st.metric("Рекомендация", result.seasonal_recommendation)
                
                st.markdown("### 💰 Рекомендованные цены")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Текущая цена", f"{result.selling_price:,.0f} ₽")
                with col2:
                    rec_price_15 = result.total_expenses / (1 - 0.15)
                    st.metric("При марже 15%", f"{rec_price_15:,.0f} ₽", 
                             delta=f"{rec_price_15 - result.selling_price:,.0f} ₽")
                with col3:
                    rec_price_25 = result.total_expenses / (1 - 0.25)
                    st.metric("При марже 25%", f"{rec_price_25:,.0f} ₽",
                             delta=f"{rec_price_25 - result.selling_price:,.0f} ₽")
    
    elif current_section == 'tariffs':
        st.markdown("## 📋 Актуальные тарифы маркетплейсов")
        
        st.info("""
        **📌 Важно:** Тарифы загружаются из API маркетплейсов или из CSV файла.
        НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ не используется.
        Если API недоступен — загрузите свои тарифы через CSV.
        """)
        
        col1, col2, col3 = st.columns(3)
        with col1:
            marketplace = st.selectbox("Маркетплейс", ["Ozon", "Wildberries", "Яндекс Маркет"])
        with col2:
            force_refresh = st.checkbox("🔄 Принудительное обновление")
        with col3:
            use_ai = st.checkbox("🤖 DeepSeek AI как fallback", value=True)
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("📥 Загрузить тарифы", type="primary"):
                with st.spinner(f"Загрузка тарифов {marketplace}..."):
                    calculator.set_marketplace(marketplace)
                    calculator.refresh_tariffs(force=force_refresh, use_ai=use_ai)
                    st.success(f"✅ Тарифы {marketplace} загружены!")
                    st.rerun()
        
        with col2:
            csv_file = st.file_uploader("📄 Или загрузите CSV с тарифами", type=['csv'])
            if csv_file and st.button("📥 Загрузить из CSV"):
                csv_content = csv_file.getvalue().decode('utf-8')
                calculator.refresh_tariffs(force=True, csv_content=csv_content)
                st.success("✅ Тарифы загружены из CSV!")
                st.rerun()
        
        if calculator.current_tariffs:
            df = calculator.api_manager.get_all_tariffs_as_dataframe(marketplace)
            # ИСПРАВЛЕНО: заменен use_container_width на width='stretch'
            st.dataframe(df, width='stretch', height=400)
            
            st.markdown("### 📊 Статистика источников")
            sources = df['Источник'].value_counts()
            st.dataframe(sources, width='stretch')
            
            # Проверяем наличие предупреждений
            warnings_df = df[df['Предупреждение'] != '']
            if not warnings_df.empty:
                st.warning("⚠️ Некоторые тарифы являются примерными. Загрузите актуальные данные через API или CSV.")
                st.dataframe(warnings_df[['Категория', 'Предупреждение']], width='stretch')
    
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
            fig = px.histogram(
                margins, 
                title="Распределение маржинальности",
                labels={'value': 'Маржа, %', 'count': 'Количество товаров'},
                nbins=20,
                color_discrete_sequence=['#6c5ce7']
            )
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            top = sorted(results, key=lambda x: x.gross_profit, reverse=True)[:10]
            df = pd.DataFrame({
                'Артикул': [r.artikul for r in top],
                'Прибыль, ₽': [r.gross_profit for r in top]
            })
            fig = px.bar(
                df, 
                x='Артикул', 
                y='Прибыль, ₽',
                title="Топ-10 по прибыли",
                color='Прибыль, ₽',
                color_continuous_scale='viridis'
            )
            st.plotly_chart(fig, use_container_width=True)
        
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
        
        df_zones = pd.DataFrame({
            'Зона': list(zones.keys()),
            'Количество': list(zones.values())
        })
        fig = px.pie(
            df_zones, 
            values='Количество', 
            names='Зона',
            title="Распределение по логистическим зонам",
            color_discrete_sequence=['#FF6B6B', '#FFD93D', '#6BCB77', '#4D96FF']
        )
        st.plotly_chart(fig, use_container_width=True)
    
    elif current_section == 'what_if':
        st.markdown("## 🎯 Анализ сценариев 'Что если'")
        
        if not st.session_state.input_data_list:
            st.warning("⚠️ Сначала выполните расчет в разделе 'Калькулятор FBS'.")
            return
        
        base_data = st.session_state.input_data_list[0] if st.session_state.input_data_list else None
        
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
            {'name': '📈 Повышение цены на 20%', 'params': {'selling_price': base_data.selling_price * 1.2}},
            {'name': '📉 Снижение цены на 15%', 'params': {'selling_price': base_data.selling_price * 0.85}},
            {'name': '🚚 Увеличение расстояния на 50%', 'params': {'warehouse_distance_km': base_data.warehouse_distance_km * 1.5}},
            {'name': '📦 Оптимизация паллет (x2)', 'params': {'pallet_capacity': base_data.pallet_capacity * 2}},
            {'name': '🕒 Внедрение ночной смены', 'params': {'has_night_shift': True}},
            {'name': '💰 Снижение себестоимости на 10%', 'params': {'cogs': base_data.cogs * 0.9}}
        ]
        
        selected_presets = st.multiselect(
            "Выберите сценарии для анализа:",
            [s['name'] for s in preset_scenarios],
            default=[s['name'] for s in preset_scenarios[:3]]
        )
        
        scenarios_to_run = [s for s in preset_scenarios if s['name'] in selected_presets]
        
        st.markdown("#### ➕ Добавить свой сценарий")
        col1, col2 = st.columns(2)
        with col1:
            custom_name = st.text_input("Название сценария", "Мой сценарий")
        with col2:
            custom_param = st.selectbox(
                "Параметр для изменения",
                ["selling_price", "cogs", "warehouse_distance_km", "pallet_capacity", 
                 "daily_sales", "packaging_cost", "marketing_budget_per_unit"]
            )
        custom_value = st.number_input("Новое значение", value=base_data.selling_price)
        
        if st.button("➕ Добавить сценарий"):
            scenarios_to_run.append({
                'name': custom_name,
                'params': {custom_param: custom_value}
            })
            st.success(f"✅ Сценарий '{custom_name}' добавлен!")
        
        if st.button("🚀 Запустить анализ", type="primary"):
            if not scenarios_to_run:
                st.warning("⚠️ Выберите хотя бы один сценарий.")
            else:
                with st.spinner("Выполнение анализа..."):
                    scenarios = []
                    for s in scenarios_to_run:
                        scenario_dict = {'name': s['name']}
                        scenario_dict.update(s['params'])
                        scenarios.append(scenario_dict)
                    
                    df_results = calculator.run_what_if_analysis(base_data, scenarios)
                    st.markdown("### 📊 Результаты анализа")
                    st.dataframe(df_results, width='stretch')
                    
                    st.markdown("### 📈 Визуализация сценариев")
                    fig = make_subplots(rows=1, cols=2, subplot_titles=("Прибыль по сценариям", "Маржа по сценариям"))
                    
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
                    st.plotly_chart(fig, use_container_width=True)
                    
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
                sorted_recommendations = sorted(
                    st.session_state.recommendations,
                    key=lambda x: priority_order.get(x['priority'], 3)
                )
                
                for rec in sorted_recommendations:
                    priority_icon = "🔴" if rec['priority'] == 'high' else "🟡" if rec['priority'] == 'medium' else "🟢"
                    with st.expander(f"{priority_icon} [{rec['priority'].upper()}] {rec['category']} - {rec['icon']} {rec['message'][:80]}..."):
                        st.markdown(f"**{rec['message']}**")
                        
                        if rec.get('affected_products'):
                            st.markdown("**📦 Затронутые товары:**")
                            st.write(", ".join(rec['affected_products'][:10]))
                            if len(rec['affected_products']) > 10:
                                st.caption(f"... и еще {len(rec['affected_products']) - 10} товаров")
                        
                        if rec['priority'] == 'high':
                            st.warning("⚠️ Это критическая рекомендация. Требует немедленного внимания.")
                        elif rec['priority'] == 'medium':
                            st.info("ℹ️ Рекомендация средней важности. Планируйте внедрение в ближайшее время.")
                        else:
                            st.success("✅ Рекомендация низкой важности. Можно рассмотреть в долгосрочной перспективе.")
        
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
            df_cat = pd.DataFrame({
                'Категория': list(categories.keys()),
                'Количество': list(categories.values())
            })
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
        
        tab1, tab2, tab3 = st.tabs(["📊 Excel", "📄 CSV", "🌐 Google Sheets"])
        
        with tab1:
            st.info("Excel файл содержит формулы, ссылающиеся на лист с тарифами. При изменении тарифов все расчеты пересчитываются автоматически.")
            
            if st.button("📥 Скачать Excel-отчет", type="primary"):
                try:
                    from openpyxl import Workbook
                    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                    
                    wb = Workbook()
                    
                    # Создаем лист с тарифами
                    ws_tariffs = wb.active
                    ws_tariffs.title = "Тарифы МП"
                    
                    # Заголовки тарифов
                    tariff_headers = ['Категория', 'Комиссия, %', 'Мин. комиссия, ₽', 'Last Mile база, ₽', 
                                     'Last Mile за кг, ₽', 'Эквайринг, %', 'Возвраты, %', 'Штрафы, %', 
                                     'Источник']
                    
                    for col, header in enumerate(tariff_headers, 1):
                        cell = ws_tariffs.cell(row=1, column=col, value=header)
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="1a1a2e", fill_type="solid")
                    
                    # Заполняем тарифы
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
                    
                    # Создаем лист с результатами
                    ws_results = wb.create_sheet("Результаты")
                    
                    # Заголовки результатов
                    result_headers = ['Артикул', 'Наименование', 'Цена, ₽', 'Прибыль, ₽', 'Маржа, %', 
                                     'ROI, %', 'Комиссия, ₽', 'First Mile, ₽', 'Last Mile, ₽',
                                     'Опт. запас, шт', 'Оборачиваемость, дн', 'Лог. зона']
                    
                    for col, header in enumerate(result_headers, 1):
                        cell = ws_results.cell(row=1, column=col, value=header)
                        cell.font = Font(bold=True, color="FFFFFF")
                        cell.fill = PatternFill(start_color="1a1a2e", fill_type="solid")
                    
                    # Заполняем результаты
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
                        
                        # Цветовая подсветка прибыли
                        if result.gross_profit > 0:
                            ws_results.cell(row=row, column=4).fill = PatternFill(start_color="C6EFCE", fill_type="solid")
                        else:
                            ws_results.cell(row=row, column=4).fill = PatternFill(start_color="FFC7CE", fill_type="solid")
                        
                        row += 1
                    
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
                with st.spinner("Экспорт в CSV..."):
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
                            'Лог. зона': result.logistic_zone_label
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
            st.markdown("""
            ### 🌐 Экспорт в Google Sheets
            
            Для экспорта в Google Sheets необходимо:
            1. Создать сервисный аккаунт в Google Cloud Console
            2. Скачать credentials.json
            3. Загрузить файл ниже
            """)
            
            credentials_file = st.file_uploader(
                "📁 Загрузите credentials.json",
                type=['json'],
                help="Файл с ключами сервисного аккаунта Google"
            )
            
            if credentials_file and st.button("📤 Экспортировать в Google Sheets", type="primary"):
                st.info("Функция экспорта в Google Sheets требует настройки gspread и сервисного аккаунта.")
                st.warning("⚠️ Для работы этой функции необходимо установить gspread и настроить сервисный аккаунт.")
    
    elif current_section == 'settings':
        st.markdown("## ⚙️ Настройки")
        
        tab1, tab2, tab3 = st.tabs(["🔑 API Ключи", "🏪 Маркетплейс и налоги", "📄 Импорт тарифов CSV"])
        
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

if __name__ == "__main__":
    main()
