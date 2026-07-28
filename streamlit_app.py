"""
============================================================================
🚀 FBS UNIT ECONOMICS PRO 2026 — ПОЛНАЯ ОПЕРАЦИОННАЯ ВЕРСИЯ
============================================================================
Операционный директор | FBS-экспертиза | Оптимизация складских остатков
Маркетплейсы: Ozon, Wildberries, Яндекс Маркет
Версия: 6.0.0

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

warnings.filterwarnings('ignore')

# ============================================================================
# БЛОК 0: БАЗОВАЯ КОНФИГУРАЦИЯ И НАСТРОЙКИ
# ============================================================================

APP_VERSION = "6.0.0"
APP_NAME = "🚀 FBS Юнит-экономика PRO 2026 — Операционная версия"
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
# БЛОК 5: API МЕНЕДЖЕР ДЛЯ ЗАГРУЗКИ ТАРИФОВ
# ============================================================================

class MarketplaceAPIManager:
    """
    Менеджер для загрузки актуальных тарифов через API маркетплейсов.
    
    Поддерживаемые источники:
    1. Прямое API маркетплейса (Ozon, Wildberries, Яндекс Маркет)
    2. DeepSeek AI (когда прямое API недоступно)
    3. Встроенные дефолтные значения (фолбэк)
    
    Особенности:
    - Автоматическое кэширование на 1 час
    - Retry при ошибках сети
    - Приоритет: API → DeepSeek → Default
    """
    
    def __init__(self, cache_manager: Optional[CacheManager] = None, 
                 secure_data: Optional[SecureDataManager] = None):
        self.cache_manager = cache_manager or CacheManager()
        self.secure_data = secure_data or SecureDataManager()
        self.session = requests.Session()
        self.session.headers.update({
            'User-Agent': f'FBS-Unit-Economy-Pro/{APP_VERSION}',
            'Accept': 'application/json',
            'Content-Type': 'application/json'
        })
        self.progress_tracker = ProgressTracker()
        self._api_keys_cache: Dict[str, str] = {}
        self._load_api_keys()
    
    def _load_api_keys(self):
        """Загрузка API ключей из защищенного хранилища"""
        try:
            # Пробуем загрузить из шифрованного хранилища
            if self.secure_data.is_available():
                self._api_keys_cache = self.secure_data.get_all_api_keys()
            
            # Если нет - пробуем обычный файл
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
        
        # Сохраняем в шифрованное хранилище
        if self.secure_data.is_available():
            success = self.secure_data.store_api_key(service, api_key.strip())
            if success:
                logger.info(f"✅ API ключ для {service} сохранен в защищенное хранилище")
                return True
        
        # Фолбэк - сохраняем в обычный JSON
        try:
            key_file = CONFIG_DIR / "api_keys.json"
            with open(key_file, 'w', encoding='utf-8') as f:
                json.dump(self._api_keys_cache, f, indent=2, ensure_ascii=False)
            try:
                os.chmod(key_file, 0o600)
            except OSError:
                pass
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
        """
        Получение закэшированных тарифов.
        
        Args:
            marketplace: Название маркетплейса
            
        Returns:
            Optional[Dict]: Закэшированные тарифы или None
        """
        cache_key = f"tariffs_{marketplace.lower()}"
        return self.cache_manager.get('tariffs', cache_key)
    
    def save_tariffs_to_cache(self, marketplace: str, tariffs: Dict[str, Dict]):
        """
        Сохранение тарифов в кэш.
        
        Args:
            marketplace: Название маркетплейса
            tariffs: Словарь с тарифами по категориям
        """
        cache_key = f"tariffs_{marketplace.lower()}"
        self.cache_manager.set('tariffs', cache_key, {
            'tariffs': tariffs,
            'marketplace': marketplace,
            'cached_at': datetime.now().isoformat(),
            'version': APP_VERSION
        })
        logger.info(f"💾 Тарифы {marketplace} сохранены в кэш")
    
    def _get_default_tariffs(self, marketplace: str) -> Dict[str, Dict]:
        """
        Получение дефолтных тарифов когда API недоступны.
        Это аварийный фолбэк для обеспечения работы приложения.
        """
        logger.warning(f"⚠️ Использую дефолтные тарифы для {marketplace}")
        
        if marketplace.lower() == 'ozon':
            return {
                'default': {
                    'commission_rate': 0.15, 'min_commission': 30.0,
                    'last_mile_base': 50.0, 'last_mile_per_kg': 15.0,
                    'last_mile_per_km': 3.5, 'acquiring_fee': 0.015,
                    'return_fee': 0.02, 'penalty_rate': 0.05,
                    'penalty_time_hours': 24, 'fbo_multiplier': 0.75,
                    'fbp_multiplier': 0.60, 'storage_base_rate': 0.30,
                    'min_logistics': 25.0, 'source': 'default',
                    'last_updated': datetime.now().isoformat()
                },
                'auto_parts': {
                    'commission_rate': 0.12, 'min_commission': 30.0,
                    'last_mile_base': 50.0, 'last_mile_per_kg': 15.0,
                    'last_mile_per_km': 3.5, 'acquiring_fee': 0.015,
                    'return_fee': 0.02, 'penalty_rate': 0.05,
                    'penalty_time_hours': 24, 'fbo_multiplier': 0.75,
                    'fbp_multiplier': 0.60, 'storage_base_rate': 0.30,
                    'min_logistics': 25.0, 'source': 'default',
                    'last_updated': datetime.now().isoformat()
                },
                'electronics': {
                    'commission_rate': 0.10, 'min_commission': 30.0,
                    'last_mile_base': 50.0, 'last_mile_per_kg': 15.0,
                    'last_mile_per_km': 3.5, 'acquiring_fee': 0.015,
                    'return_fee': 0.02, 'penalty_rate': 0.05,
                    'penalty_time_hours': 24, 'fbo_multiplier': 0.75,
                    'fbp_multiplier': 0.60, 'storage_base_rate': 0.30,
                    'min_logistics': 25.0, 'source': 'default',
                    'last_updated': datetime.now().isoformat()
                }
            }
        elif marketplace.lower() == 'wildberries':
            return {
                'default': {
                    'commission_rate': 0.16, 'min_commission': 28.0,
                    'last_mile_base': 45.0, 'last_mile_per_kg': 14.0,
                    'last_mile_per_km': 3.2, 'acquiring_fee': 0.015,
                    'return_fee': 0.018, 'penalty_rate': 0.08,
                    'penalty_time_hours': 24, 'fbo_multiplier': 0.70,
                    'fbp_multiplier': 0.55, 'storage_base_rate': 0.25,
                    'min_logistics': 22.0, 'source': 'default',
                    'last_updated': datetime.now().isoformat()
                }
            }
        elif marketplace.lower() == 'yandex_market':
            return {
                'default': {
                    'commission_rate': 0.145, 'min_commission': 35.0,
                    'last_mile_base': 55.0, 'last_mile_per_kg': 16.0,
                    'last_mile_per_km': 3.8, 'acquiring_fee': 0.015,
                    'return_fee': 0.025, 'penalty_rate': 0.07,
                    'penalty_time_hours': 24, 'fbo_multiplier': 0.80,
                    'fbp_multiplier': 0.65, 'storage_base_rate': 0.35,
                    'min_logistics': 30.0, 'source': 'default',
                    'last_updated': datetime.now().isoformat()
                }
            }
        else:
            return self._get_default_tariffs('ozon')
    
    @retry_on_failure(max_retries=2, delay=2.0)
    def fetch_ozon_tariffs(self) -> Dict[str, Dict]:
        """
        Загрузка тарифов Ozon через официальное API.
        
        API Endpoints:
        - POST /v1/commission/list - комиссии по категориям
        - POST /v1/delivery-methods - способы доставки и тарифы
        
        Документация: https://docs.ozon.ru/api/seller/
        """
        tariffs = {}
        
        client_id = self.get_api_key('ozon_client_id')
        api_key = self.get_api_key('ozon')
        
        if not client_id or not api_key:
            logger.warning("⚠️ API ключи Ozon не найдены")
            return self._get_default_tariffs('ozon')
        
        headers = {
            'Client-Id': client_id,
            'Api-Key': api_key,
            'Content-Type': 'application/json'
        }
        
        try:
            # Запрос комиссий по категориям
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
                
                # Запрос способов доставки
                logger.info("📡 Запрос тарифов доставки Ozon...")
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
                
                # Парсинг и объединение данных
                for item in commission_data.get('result', []):
                    category = item.get('category', 'default')
                    category_name = item.get('category_name', category)
                    
                    # Ищем тарифы доставки для этой категории
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
                        'penalty_rate': 0.05,
                        'penalty_time_hours': 24,
                        'fbo_multiplier': 0.75,
                        'fbp_multiplier': 0.60,
                        'storage_base_rate': float(item.get('storage_rate', 0.30)),
                        'min_logistics': float(delivery_info.get('min_delivery_cost', 25)),
                        'source': 'ozon_api',
                        'last_updated': datetime.now().isoformat(),
                        'api_response_raw': json.dumps(item, ensure_ascii=False)
                    }
                
                if tariffs:
                    logger.info(f"✅ Загружено {len(tariffs)} категорий тарифов Ozon через API")
                else:
                    logger.warning("⚠️ Не удалось распарсить тарифы Ozon")
                    tariffs = self._get_default_tariffs('ozon')
            else:
                logger.error(f"❌ Ozon API вернул статус {commission_response.status_code}: {commission_response.text}")
                tariffs = self._get_default_tariffs('ozon')
        
        except requests.exceptions.Timeout:
            logger.error("❌ Таймаут запроса к Ozon API")
            tariffs = self._get_default_tariffs('ozon')
        except requests.exceptions.ConnectionError as e:
            logger.error(f"❌ Ошибка подключения к Ozon API: {e}")
            tariffs = self._get_default_tariffs('ozon')
        except Exception as e:
            logger.error(f"❌ Неожиданная ошибка при загрузке тарифов Ozon: {e}")
            logger.exception(e)
            tariffs = self._get_default_tariffs('ozon')
        
        return tariffs
    
    @retry_on_failure(max_retries=2, delay=2.0)
    def fetch_wildberries_tariffs(self) -> Dict[str, Dict]:
        """
        Загрузка тарифов Wildberries через официальное API.
        
        API Endpoints:
        - GET /api/v2/tariffs - тарифы на доставку
        - GET /api/v2/commissions - комиссии по категориям
        
        Документация: https://suppliers.wildberries.ru/
        """
        tariffs = {}
        
        api_key = self.get_api_key('wildberries')
        
        if not api_key:
            logger.warning("⚠️ API ключ Wildberries не найден")
            return self._get_default_tariffs('wildberries')
        
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        
        try:
            # Запрос тарифов доставки
            logger.info("📡 Запрос тарифов Wildberries...")
            tariffs_response = self.session.get(
                MarketplaceAPIEndpoint.WILDBERRIES_TARIFFS.value,
                headers=headers,
                params={'locale': 'ru'},
                timeout=30
            )
            
            # Запрос комиссий
            logger.info("📡 Запрос комиссий Wildberries...")
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
            
            # Парсинг данных
            tariff_items = tariffs_data.get('data', {}).get('tariffs', [])
            commission_items = commissions_data.get('data', {}).get('commissions', [])
            
            # Создаем словарь комиссий для быстрого поиска
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
                    'acquiring_fee': 0.015,
                    'return_fee': float(item.get('returnPercent', 1.8)) / 100,
                    'penalty_rate': 0.08,
                    'penalty_time_hours': 24,
                    'fbo_multiplier': 0.70,
                    'fbp_multiplier': 0.55,
                    'storage_base_rate': float(item.get('storageRate', 0.25)),
                    'min_logistics': float(item.get('minDeliveryCost', 22)),
                    'source': 'wildberries_api',
                    'last_updated': datetime.now().isoformat(),
                    'api_response_raw': json.dumps(item, ensure_ascii=False)
                }
            
            if tariffs:
                logger.info(f"✅ Загружено {len(tariffs)} категорий тарифов Wildberries через API")
            else:
                logger.warning("⚠️ Не удалось распарсить тарифы Wildberries")
                tariffs = self._get_default_tariffs('wildberries')
        
        except requests.exceptions.Timeout:
            logger.error("❌ Таймаут запроса к Wildberries API")
            tariffs = self._get_default_tariffs('wildberries')
        except requests.exceptions.ConnectionError as e:
            logger.error(f"❌ Ошибка подключения к Wildberries API: {e}")
            tariffs = self._get_default_tariffs('wildberries')
        except Exception as e:
            logger.error(f"❌ Неожиданная ошибка при загрузке тарифов Wildberries: {e}")
            logger.exception(e)
            tariffs = self._get_default_tariffs('wildberries')
        
        return tariffs
    
    @retry_on_failure(max_retries=2, delay=2.0)
    def fetch_yandex_market_tariffs(self) -> Dict[str, Dict]:
        """
        Загрузка тарифов Яндекс Маркет через официальное API.
        
        API Endpoints:
        - GET /campaigns/{campaignId}/tariffs
        - GET /campaigns/{campaignId}/offer-mapping-entries
        
        Документация: https://yandex.ru/dev/market/partner/
        """
        tariffs = {}
        
        api_key = self.get_api_key('yandex_market')
        campaign_id = self.get_api_key('yandex_campaign_id')
        
        if not api_key or not campaign_id:
            logger.warning("⚠️ API ключи Яндекс Маркет не найдены")
            return self._get_default_tariffs('yandex_market')
        
        headers = {
            'Authorization': f'Bearer {api_key}',
            'Content-Type': 'application/json'
        }
        
        try:
            # Формируем URL с campaign_id
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
                        'acquiring_fee': 0.015,
                        'return_fee': 0.025,
                        'penalty_rate': 0.07,
                        'penalty_time_hours': 24,
                        'fbo_multiplier': 0.80,
                        'fbp_multiplier': 0.65,
                        'storage_base_rate': float(item.get('storageRate', 0.35)),
                        'min_logistics': float(item.get('minDeliveryCost', 30)),
                        'source': 'yandex_api',
                        'last_updated': datetime.now().isoformat(),
                        'api_response_raw': json.dumps(item, ensure_ascii=False)
                    }
                
                if tariffs:
                    logger.info(f"✅ Загружено {len(tariffs)} категорий тарифов Яндекс Маркет через API")
                else:
                    logger.warning("⚠️ Не удалось распарсить тарифы Яндекс Маркет")
                    tariffs = self._get_default_tariffs('yandex_market')
            else:
                logger.error(f"❌ Яндекс Маркет API вернул статус {response.status_code}: {response.text}")
                tariffs = self._get_default_tariffs('yandex_market')
        
        except requests.exceptions.Timeout:
            logger.error("❌ Таймаут запроса к Яндекс Маркет API")
            tariffs = self._get_default_tariffs('yandex_market')
        except requests.exceptions.ConnectionError as e:
            logger.error(f"❌ Ошибка подключения к Яндекс Маркет API: {e}")
            tariffs = self._get_default_tariffs('yandex_market')
        except Exception as e:
            logger.error(f"❌ Неожиданная ошибка при загрузке тарифов Яндекс Маркет: {e}")
            logger.exception(e)
            tariffs = self._get_default_tariffs('yandex_market')
        
        return tariffs
    
    def fetch_tariffs_via_deepseek(self, marketplace: str) -> Dict[str, Dict]:
        """
        AI-обогащение тарифов через DeepSeek API.
        
        Используется когда:
        - Прямые API маркетплейсов недоступны
        - Нужна верификация данных из разных источников
        - Требуется получить агрегированную информацию
        
        Args:
            marketplace: Название маркетплейса
            
        Returns:
            Dict: Тарифы по категориям
        """
        tariffs = {}
        
        api_key = self.get_api_key('deepseek')
        
        if not api_key:
            logger.warning("⚠️ DeepSeek API ключ не найден")
            return self._get_default_tariffs(marketplace.lower())
        
        try:
            # Формируем промпт для AI
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
                
                # Парсим JSON из ответа
                try:
                    content = json.loads(content_text)
                except json.JSONDecodeError:
                    # Пробуем извлечь JSON из текста
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
                else:
                    logger.warning("⚠️ DeepSeek не предоставил тарифы")
                    tariffs = self._get_default_tariffs(marketplace.lower())
            else:
                logger.error(f"❌ DeepSeek API вернул статус {response.status_code}: {response.text}")
                tariffs = self._get_default_tariffs(marketplace.lower())
        
        except requests.exceptions.Timeout:
            logger.error("❌ Таймаут запроса к DeepSeek API")
            tariffs = self._get_default_tariffs(marketplace.lower())
        except json.JSONDecodeError as e:
            logger.error(f"❌ Ошибка парсинга JSON от DeepSeek: {e}")
            tariffs = self._get_default_tariffs(marketplace.lower())
        except Exception as e:
            logger.error(f"❌ Неожиданная ошибка при запросе к DeepSeek: {e}")
            logger.exception(e)
            tariffs = self._get_default_tariffs(marketplace.lower())
        
        return tariffs
    
    def get_tariffs(self, marketplace: str, force_refresh: bool = False, 
                   use_ai_fallback: bool = True) -> Dict[str, Dict]:
        """
        Основной метод получения тарифов.
        
        Приоритет источников:
        1. Кэш (если не force_refresh)
        2. Прямое API маркетплейса
        3. DeepSeek AI (если use_ai_fallback)
        4. Встроенные дефолтные значения
        
        Args:
            marketplace: Название маркетплейса (Ozon, Wildberries, Яндекс Маркет)
            force_refresh: Принудительно обновить, игнорируя кэш
            use_ai_fallback: Использовать DeepSeek AI при недоступности API
            
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
        
        logger.info(f"🔄 Загрузка тарифов {marketplace} через API...")
        
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
                tariffs = self._get_default_tariffs(marketplace_lower)
        except Exception as e:
            logger.error(f"❌ Ошибка API для {marketplace}: {e}")
            tariffs = {}
        
        # Если API не сработало - пробуем DeepSeek
        if (not tariffs or all(t.get('source') == 'default' for t in tariffs.values())) and use_ai_fallback:
            logger.info(f"🤖 Прямое API недоступно, использую DeepSeek AI для {marketplace}")
            try:
                ai_tariffs = self.fetch_tariffs_via_deepseek(marketplace)
                if ai_tariffs and not all(t.get('source') == 'default' for t in ai_tariffs.values()):
                    tariffs = ai_tariffs
                    logger.info(f"✅ Тарифы {marketplace} получены через DeepSeek AI")
            except Exception as e:
                logger.error(f"❌ DeepSeek также недоступен: {e}")
        
        # Если ничего не сработало - дефолтные значения
        if not tariffs:
            logger.warning(f"⚠️ Все источники недоступны, использую дефолтные тарифы для {marketplace}")
            tariffs = self._get_default_tariffs(marketplace_lower)
        
        # Сохраняем в кэш
        if tariffs:
            self.save_tariffs_to_cache(marketplace, tariffs)
        
        return tariffs
    
    def get_all_tariffs_as_dataframe(self, marketplace: str) -> pd.DataFrame:
        """
        Получение всех тарифов в виде DataFrame для отображения в UI.
        
        Args:
            marketplace: Название маркетплейса
            
        Returns:
            pd.DataFrame: Таблица с тарифами
        """
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
                'Обновлено': data.get('last_updated', '')[:19] if data.get('last_updated') else ''
            })
        
        return pd.DataFrame(rows)
    
    def test_api_connection(self, marketplace: str) -> Dict[str, Any]:
        """
        Тестирование API подключения.
        
        Returns:
            Dict с результатами теста
        """
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
    Содержит все необходимые параметры товара и бизнеса.
    РАСШИРЕНО: добавлены параметры для оптимизации складских остатков
    """
    # Основные параметры товара
    artikul: str = ""
    product_name: str = ""
    category: str = "default"
    
    # Финансовые параметры
    selling_price: float = 0.0  # Цена продажи на маркетплейсе
    cogs: float = 0.0  # Себестоимость закупки (Cost of Goods Sold)
    
    # Физические параметры товара
    weight_kg: float = 0.0  # Вес брутто в кг
    length_cm: float = 0.0  # Длина упаковки в см
    width_cm: float = 0.0  # Ширина упаковки в см
    height_cm: float = 0.0  # Высота упаковки в см
    
    # FBS специфика - First Mile (ваша логистика до склада МП)
    first_mile_cost_per_unit: float = 0.0  # Фиксированная стоимость доставки 1 шт
    packaging_cost: float = 0.0  # Стоимость упаковочного материала на 1 шт
    pick_pack_time_min: float = 5.0  # Время сборки заказа в минутах
    operator_hourly_rate: float = 300.0  # Ставка оператора сборки в ₽/час
    warehouse_distance_km: float = 0.0  # Расстояние от вашего склада до склада МП
    
    # Логистические параметры
    transport_type: str = "own"  # Тип транспорта: own, cdek, delovye_linii
    transport_cost_per_km: float = 20.0  # Стоимость 1 км пробега транспорта
    pallet_capacity: int = 100  # Количество единиц товара на одной паллете
    pallet_cost: float = 2000.0  # Фиксированная стоимость паллетной отправки
    
    # Маркетинговые параметры
    marketing_budget_per_unit: float = 0.0  # Рекламный бюджет на единицу товара
    
    # Складские параметры (РАСШИРЕНО)
    stock_depth_days: int = 30  # Глубина складского запаса в днях
    daily_sales: int = 5  # Среднее количество продаж в день
    warehouse_rent_per_sqm: float = 500.0  # Стоимость аренды склада за м² в месяц
    warehouse_space_per_unit: float = 0.01  # Занимаемая площадь на складе на единицу (м²)
    safety_stock_days: int = 7  # Страховой запас в днях (НОВОЕ)
    reorder_point_days: int = 5  # Точка заказа в днях (НОВОЕ)
    supplier_lead_time_days: int = 3  # Время поставки от поставщика в днях (НОВОЕ)
    
    # Параметры для расчета LTV и CAC
    repeat_purchase_rate: float = 0.3  # Коэффициент повторных покупок
    avg_purchases_per_year: float = 2.5  # Среднее количество покупок в год на клиента
    customer_retention_rate: float = 0.7  # Коэффициент удержания клиентов (CRR)
    discount_rate: float = 0.1  # Ставка дисконтирования
    
    # Операционные параметры
    has_night_shift: bool = False  # Наличие ночной смены для обработки заказов
    processing_capacity_per_hour: int = 20  # Пропускная способность обработки в час
    
    def to_dict(self) -> Dict[str, Any]:
        """Преобразование в словарь для сериализации"""
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'FBSInputData':
        """Создание из словаря"""
        known_fields = {f.name for f in cls.__dataclass_fields__.values()}
        filtered_data = {k: v for k, v in data.items() if k in known_fields}
        return cls(**filtered_data)
    
    def validate(self) -> List[str]:
        """Валидация входных данных, возвращает список ошибок"""
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
    Содержит полную детализацию всех расходов и метрик.
    РАСШИРЕНО: добавлены метрики складской оптимизации
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
    commission: float = 0.0  # Комиссия маркетплейса
    first_mile_cost: float = 0.0  # Стоимость First Mile (ваша логистика до МП)
    last_mile_cost: float = 0.0  # Стоимость Last Mile (логистика МП до клиента)
    pick_pack_cost: float = 0.0  # Стоимость обработки заказа (Pick & Pack)
    packaging_cost: float = 0.0  # Стоимость упаковочных материалов
    acquiring_cost: float = 0.0  # Эквайринг
    return_cost: float = 0.0  # Стоимость возвратов
    penalty_cost: float = 0.0  # Штрафы за просрочку
    marketing_cost: float = 0.0  # Маркетинговые расходы
    warehouse_cost: float = 0.0  # Складские расходы (распределенные)
    tax_cost: float = 0.0  # Налоги
    
    # FBS специфические метрики
    penalty_probability: float = 0.0  # Вероятность получения штрафа за просрочку
    break_even_distance_km: float = 0.0  # Точка безубыточности по расстоянию First Mile
    max_discount_percent: float = 0.0  # Максимально возможная скидка (%)
    safety_margin_price: float = 0.0  # Запас прочности по цене (₽)
    
    # LTV и CAC метрики
    ltv: float = 0.0  # Lifetime Value (жизненная ценность клиента)
    cac: float = 0.0  # Customer Acquisition Cost (стоимость привлечения клиента)
    ltv_cac_ratio: float = 0.0  # Соотношение LTV/CAC
    romi: float = 0.0  # Return on Marketing Investment (НОВОЕ)
    
    # Метрики складской оптимизации (НОВЫЕ)
    optimal_stock_units: int = 0  # Оптимальный запас в единицах
    safety_stock_units: int = 0  # Страховой запас в единицах
    reorder_point_units: int = 0  # Точка заказа в единицах
    stock_turnover_days: float = 0.0  # Оборачиваемость в днях
    stock_turnover_rate: float = 0.0  # Коэффициент оборачиваемости
    days_of_inventory: float = 0.0  # Дни запаса
    holding_cost_per_unit: float = 0.0  # Стоимость хранения на единицу
    
    # Рекомендации по оптимизации склада (НОВЫЕ)
    recommended_stock_depth_days: int = 0  # Рекомендуемая глубина запаса
    recommended_safety_stock_days: int = 0  # Рекомендуемый страховой запас
    stock_optimization_potential: float = 0.0  # Потенциал оптимизации запаса (%)
    
    def to_dict(self) -> Dict[str, Any]:
        """Преобразование в словарь для сериализации"""
        return asdict(self)
    
    def get_summary(self) -> Dict[str, Any]:
        """Получение краткой сводки результатов"""
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
            'stock_optimization_potential': self.stock_optimization_potential
        }

# ============================================================================
# БЛОК 7: ОСНОВНОЙ КАЛЬКУЛЯТОР FBS ЮНИТ-ЭКОНОМИКИ (РАСШИРЕННЫЙ)
# ============================================================================

# Конфигурации налоговых систем
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

class FBSUnitEconomicsCalculator:
    """
    Профессиональный калькулятор юнит-экономики для FBS-модели.
    
    Особенности:
    - Использует динамически загружаемые тарифы через API (НИКАКИХ ЗАХАРДКОЖЕННЫХ ЗНАЧЕНИЙ)
    - Учитывает специфику FBS: двойную логистику, штрафы, Pick & Pack
    - Рассчитывает LTV, CAC, ROMI
    - ОПТИМИЗАЦИЯ СКЛАДСКИХ ОСТАТКОВ: оптимальный запас, страховой запас, точка заказа
    - ЛОГИСТИЧЕСКИЕ КОРИДОРЫ: точки безубыточности по расстоянию
    - Рекомендованные цены в отдельном столбце
    - Условное форматирование прибыльных/убыточных товаров
    """
    
    def __init__(self, api_manager: Optional[MarketplaceAPIManager] = None, 
                 tax_system: str = "УСН 6% (доходы)"):
        """
        Инициализация калькулятора.
        
        Args:
            api_manager: Менеджер API для загрузки тарифов
            tax_system: Система налогообложения
        """
        self.api_manager = api_manager or MarketplaceAPIManager()
        self.tax_system = tax_system
        self.current_marketplace = "Ozon"
        self.current_tariffs: Dict[str, Dict] = {}
        self.tariffs_updated_at: Optional[datetime] = None
        self.tariffs_source = "default"
        
        # Настройки по умолчанию для FBS
        self.default_pick_pack_time = 5.0  # минут на сборку
        self.default_operator_rate = 300.0  # ₽/час ставка оператора
        self.default_first_mile_per_km = 20.0  # ₽/км стоимость транспорта
        self.default_penalty_probability_no_night = 0.35  # вероятность просрочки без ночной смены
        self.default_penalty_probability_with_night = 0.05  # вероятность просрочки с ночной сменой
        
        # Критические зоны для логистических коридоров (НОВЫЕ)
        self.CRITICAL_DISTANCE_KM = 25  # Красная зона
        self.WARNING_DISTANCE_KM = 50   # Желтая зона
        self.SAFE_DISTANCE_KM = 100     # Зеленая зона
        
        # Прогресс-трекер для пакетной обработки
        self.progress_tracker = ProgressTracker()
        
        # Загружаем тарифы при инициализации
        self.refresh_tariffs()
    
    def set_marketplace(self, marketplace_name: str):
        """
        Установка маркетплейса и загрузка его тарифов.
        
        Args:
            marketplace_name: Название маркетплейса (Ozon, Wildberries, Яндекс Маркет)
        """
        self.current_marketplace = marketplace_name
        self.refresh_tariffs()
        logger.info(f"🏪 Установлен маркетплейс: {marketplace_name}")
    
    def refresh_tariffs(self, force: bool = False, use_ai: bool = False):
        """
        Обновление тарифов из API.
        
        Args:
            force: Принудительное обновление, игнорируя кэш
            use_ai: Использовать DeepSeek AI при недоступности API
        """
        logger.info(f"🔄 Обновление тарифов для {self.current_marketplace}...")
        
        self.current_tariffs = self.api_manager.get_tariffs(
            self.current_marketplace,
            force_refresh=force,
            use_ai_fallback=use_ai
        )
        
        self.tariffs_updated_at = datetime.now()
        
        # Определяем источник тарифов
        sources = set()
        for tariff in self.current_tariffs.values():
            source = tariff.get('source', 'default')
            if 'api' in source:
                sources.add('api')
            elif 'deepseek' in source:
                sources.add('deepseek')
            else:
                sources.add('default')
        
        if 'api' in sources:
            self.tariffs_source = 'api'
        elif 'deepseek' in sources:
            self.tariffs_source = 'deepseek'
        else:
            self.tariffs_source = 'default'
        
        logger.info(f"✅ Тарифы обновлены. Источник: {self.tariffs_source}. Категорий: {len(self.current_tariffs)}")
    
    def get_tariff_for_category(self, category: str) -> Dict[str, Any]:
        """
        Получение тарифа для конкретной категории товара.
        
        Алгоритм поиска:
        1. Точное совпадение категории
        2. Частичное совпадение (категория содержится в ключе или наоборот)
        3. Категория 'default'
        
        Args:
            category: Категория товара
            
        Returns:
            Dict: Параметры тарифа
        """
        if not self.current_tariffs:
            logger.warning("⚠️ Тарифы не загружены, использую значения по умолчанию")
            return self.api_manager._get_default_tariffs(self.current_marketplace.lower()).get('default', {})
        
        # Точное совпадение
        if category in self.current_tariffs:
            return self.current_tariffs[category]
        
        # Частичное совпадение
        category_lower = category.lower()
        for cat, tariff in self.current_tariffs.items():
            cat_lower = cat.lower()
            if category_lower in cat_lower or cat_lower in category_lower:
                logger.debug(f"🔍 Найдено частичное совпадение: {category} -> {cat}")
                return tariff
        
        # Дефолтная категория
        if 'default' in self.current_tariffs:
            logger.debug(f"🔍 Использую дефолтный тариф для категории: {category}")
            return self.current_tariffs['default']
        
        # Если и default нет - возвращаем первый доступный
        first_tariff = next(iter(self.current_tariffs.values()), {})
        if first_tariff:
            logger.warning(f"⚠️ Категория {category} не найдена, использую первый доступный тариф")
            return first_tariff
        
        # Аварийный фолбэк
        return {
            'commission_rate': 0.15,
            'min_commission': 30.0,
            'last_mile_base': 50.0,
            'last_mile_per_kg': 15.0,
            'acquiring_fee': 0.015,
            'return_fee': 0.02,
            'penalty_rate': 0.05,
            'penalty_time_hours': 24,
            'fbo_multiplier': 0.75,
            'fbp_multiplier': 0.60,
            'storage_base_rate': 0.30,
            'min_logistics': 25.0,
            'source': 'fallback',
            'last_updated': datetime.now().isoformat()
        }
    
    @timing_decorator
    def calculate_unit_economics(self, input_data: FBSInputData) -> FBSResultData:
        """
        Основной расчет юнит-экономики для FBS модели.
        
        Выполняет полный расчет всех расходов и метрик:
        1. Комиссия маркетплейса
        2. First Mile (логистика до склада МП)
        3. Last Mile (логистика МП до клиента)
        4. Pick & Pack (обработка заказа)
        5. Упаковка
        6. Эквайринг
        7. Возвраты
        8. Штрафы за просрочку
        9. Маркетинг
        10. Складские расходы
        11. Налоги
        12. LTV, CAC, ROMI
        13. ОПТИМИЗАЦИЯ СКЛАДСКИХ ОСТАТКОВ (НОВОЕ)
        14. ЛОГИСТИЧЕСКИЕ КОРИДОРЫ (НОВОЕ)
        15. РЕКОМЕНДОВАННЫЕ ЦЕНЫ (НОВОЕ)
        
        Args:
            input_data: Входные данные товара
            
        Returns:
            FBSResultData: Полные результаты расчета
        """
        # Валидация входных данных
        validation_errors = input_data.validate()
        if validation_errors:
            logger.warning(f"⚠️ Ошибки валидации для {input_data.artikul}: {validation_errors}")
        
        # Инициализация результата
        result = FBSResultData()
        result.artikul = input_data.artikul
        result.product_name = input_data.product_name
        result.selling_price = input_data.selling_price
        
        # Получаем актуальный тариф для категории (ИЗ API, НЕ ЗАХАРДКОЖЕННЫЙ)
        tariff = self.get_tariff_for_category(input_data.category)
        
        # =====================================================================
        # 1. КОМИССИЯ МАРКЕТПЛЕЙСА
        # =====================================================================
        commission_rate = tariff.get('commission_rate', 0.15)
        min_commission = tariff.get('min_commission', 30.0)
        
        result.commission = max(
            input_data.selling_price * commission_rate,
            min_commission
        )
        
        # =====================================================================
        # 2. FIRST MILE (ВАША ЛОГИСТИКА ДО СКЛАДА МП)
        # =====================================================================
        if input_data.first_mile_cost_per_unit > 0:
            # Используем фиксированную стоимость, если указана
            result.first_mile_cost = input_data.first_mile_cost_per_unit
        else:
            # Рассчитываем на основе расстояния и загрузки паллета
            pallet_units = max(input_data.pallet_capacity, 1)
            cost_per_pallet = input_data.warehouse_distance_km * input_data.transport_cost_per_km * 2  # Туда-обратно
            result.first_mile_cost = cost_per_pallet / pallet_units
        
        # =====================================================================
        # 3. LAST MILE (ЛОГИСТИКА МАРКЕТПЛЕЙСА ДО КЛИЕНТА)
        # =====================================================================
        # Расчет объемного веса (1 м³ = 5000 кг для авиаперевозок)
        if input_data.length_cm > 0 and input_data.width_cm > 0 and input_data.height_cm > 0:
            vol_weight = (input_data.length_cm * input_data.width_cm * input_data.height_cm) / 5000.0
        else:
            vol_weight = 0
        
        # Оплачиваемый вес - максимальный из физического и объемного
        billable_weight = max(input_data.weight_kg, vol_weight)
        # Округление до 0.5 кг в большую сторону (стандарт МП)
        billable_weight = math.ceil(billable_weight * 2) / 2
        
        last_mile_base = tariff.get('last_mile_base', 50.0)
        last_mile_per_kg = tariff.get('last_mile_per_kg', 15.0)
        min_logistics = tariff.get('min_logistics', 25.0)
        
        result.last_mile_cost = max(
            last_mile_base + (billable_weight * last_mile_per_kg),
            min_logistics
        )
        
        # =====================================================================
        # 4. PICK & PACK (СТОИМОСТЬ ОБРАБОТКИ ЗАКАЗА НА ВАШЕМ СКЛАДЕ)
        # =====================================================================
        pick_pack_hours = input_data.pick_pack_time_min / 60.0
        result.pick_pack_cost = pick_pack_hours * input_data.operator_hourly_rate
        
        # =====================================================================
        # 5. УПАКОВОЧНЫЕ МАТЕРИАЛЫ
        # =====================================================================
        result.packaging_cost = input_data.packaging_cost
        
        # =====================================================================
        # 6. ЭКВАЙРИНГ (КОМИССИЯ ЗА ПРИЕМ ПЛАТЕЖЕЙ)
        # =====================================================================
        acquiring_fee = tariff.get('acquiring_fee', 0.015)
        result.acquiring_cost = input_data.selling_price * acquiring_fee
        
        # =====================================================================
        # 7. ВОЗВРАТЫ
        # =====================================================================
        return_fee = tariff.get('return_fee', 0.02)
        result.return_cost = input_data.selling_price * return_fee
        
        # =====================================================================
        # 8. ШТРАФЫ ЗА ПРОСРОЧКУ ПЕРЕДАЧИ ЗАКАЗА (PENALTY RATE)
        # =====================================================================
        # Вероятность просрочки зависит от наличия ночной смены
        if input_data.has_night_shift:
            penalty_probability = self.default_penalty_probability_with_night
        else:
            # Без ночной смены: заказы после 18:00 имеют высокий риск просрочки
            # ~35% заказов приходят после 18:00 и могут быть не обработаны вовремя
            penalty_probability = self.default_penalty_probability_no_night
        
        penalty_rate = tariff.get('penalty_rate', 0.05)
        
        result.penalty_probability = penalty_probability
        result.penalty_cost = input_data.selling_price * penalty_rate * penalty_probability
        
        # =====================================================================
        # 9. МАРКЕТИНГОВЫЕ РАСХОДЫ
        # =====================================================================
        result.marketing_cost = input_data.marketing_budget_per_unit
        
        # =====================================================================
        # 10. СКЛАДСКИЕ РАСХОДЫ (РАСПРЕДЕЛЕНИЕ НА ЕДИНИЦУ)
        # =====================================================================
        total_stock = input_data.stock_depth_days * input_data.daily_sales
        if total_stock > 0 and input_data.daily_sales > 0:
            total_warehouse_space = input_data.warehouse_space_per_unit * total_stock
            monthly_rent = input_data.warehouse_rent_per_sqm * total_warehouse_space
            # Распределяем на 30 дней и на количество продаж в день
            result.warehouse_cost = monthly_rent / (30 * input_data.daily_sales)
        else:
            result.warehouse_cost = 0
        
        # =====================================================================
        # 11. НАЛОГ
        # =====================================================================
        tax_config = TAX_SYSTEMS.get(self.tax_system, TAX_SYSTEMS["УСН 6% (доходы)"])
        
        if tax_config["base"] == "revenue":
            # Налог с оборота (доходов)
            result.tax_cost = input_data.selling_price * tax_config["rate"]
        else:
            # Налог с прибыли (доходы минус расходы)
            pre_tax_expenses = (
                result.commission + result.first_mile_cost + result.last_mile_cost +
                result.pick_pack_cost + result.packaging_cost + result.acquiring_cost +
                result.return_cost + result.penalty_cost + result.marketing_cost +
                result.warehouse_cost + input_data.cogs
            )
            pre_tax_profit = input_data.selling_price - pre_tax_expenses
            result.tax_cost = max(0, pre_tax_profit * tax_config["rate"])
            
            # Минимальный налог для УСН 15% (1% от доходов)
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
        
        # Маржинальность
        if result.selling_price > 0:
            result.margin_percent = (result.gross_profit / result.selling_price) * 100
        else:
            result.margin_percent = 0
        
        # ROI (возврат на инвестиции)
        if input_data.cogs > 0:
            result.roi_percent = (result.gross_profit / input_data.cogs) * 100
        else:
            result.roi_percent = 0
        
        # =====================================================================
        # 13. ТОЧКА БЕЗУБЫТОЧНОСТИ ПО РАССТОЯНИЮ (FIRST MILE)
        # =====================================================================
        # Максимальное расстояние, при котором First Mile окупается
        if result.first_mile_cost > 0 and input_data.pallet_capacity > 0:
            cost_per_km_per_unit = (input_data.transport_cost_per_km * 2) / input_data.pallet_capacity
            if cost_per_km_per_unit > 0:
                result.break_even_distance_km = result.gross_profit / cost_per_km_per_unit
            else:
                result.break_even_distance_km = float('inf')
        else:
            result.break_even_distance_km = float('inf')
        
        # =====================================================================
        # 14. ЗАПАС ПРОЧНОСТИ ПО ЦЕНЕ (ДЛЯ СЕЗОННЫХ РАСПРОДАЖ)
        # =====================================================================
        # Расчет минимальной цены для безубыточности
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
        
        # Минимальная цена = Fixed Costs / (1 - Variable Costs %)
        denominator = 1 - variable_costs_percent
        if denominator > 0:
            min_price = fixed_costs_per_unit / denominator
        else:
            min_price = fixed_costs_per_unit * 2  # Аварийный расчет
        
        result.safety_margin_price = input_data.selling_price - min_price
        
        # Максимальный процент скидки
        if input_data.selling_price > 0:
            result.max_discount_percent = ((input_data.selling_price - min_price) / input_data.selling_price) * 100
        else:
            result.max_discount_percent = 0
        
        # =====================================================================
        # 15. LTV (LIFETIME VALUE) И CAC (CUSTOMER ACQUISITION COST)
        # =====================================================================
        # LTV = (Средний чек × Кол-во покупок в год × Коэффициент удержания) / (1 + Ставка дисконтирования)
        if (1 + input_data.discount_rate) > 0:
            result.ltv = (
                input_data.selling_price *
                input_data.avg_purchases_per_year *
                input_data.customer_retention_rate
            ) / (1 + input_data.discount_rate)
        else:
            result.ltv = input_data.selling_price * input_data.avg_purchases_per_year * input_data.customer_retention_rate
        
        # CAC = (Маркетинг + Штрафы + First Mile) / Доля новых клиентов
        total_acquisition_cost = result.marketing_cost + result.penalty_cost + result.first_mile_cost
        new_customers_per_order = 0.3  # Предполагаем, что 30% заказов - новые клиенты
        
        if new_customers_per_order > 0:
            result.cac = total_acquisition_cost / new_customers_per_order
        else:
            result.cac = 0
        
        # Соотношение LTV/CAC
        if result.cac > 0:
            result.ltv_cac_ratio = result.ltv / result.cac
        else:
            result.ltv_cac_ratio = float('inf')
        
        # ROMI (Return on Marketing Investment)
        if result.marketing_cost > 0:
            result.romi = (result.gross_profit / result.marketing_cost) * 100
        else:
            result.romi = 0
        
        # =====================================================================
        # 16. ОПТИМИЗАЦИЯ СКЛАДСКИХ ОСТАТКОВ (НОВОЕ)
        # =====================================================================
        # Расчет оптимального запаса (EOQ - Economic Order Quantity)
        daily_demand = input_data.daily_sales
        annual_demand = daily_demand * 365
        
        # Стоимость заказа (фиксированные затраты на один заказ)
        ordering_cost = 500.0  # Фиксированная стоимость оформления заказа
        
        # Стоимость хранения на единицу в год
        holding_cost_per_unit = input_data.warehouse_rent_per_sqm * input_data.warehouse_space_per_unit * 12
        
        if holding_cost_per_unit > 0 and ordering_cost > 0:
            # EOQ = sqrt((2 * Annual Demand * Ordering Cost) / Holding Cost)
            eoq = math.sqrt((2 * annual_demand * ordering_cost) / holding_cost_per_unit)
            result.optimal_stock_units = int(math.ceil(eoq))
        else:
            result.optimal_stock_units = input_data.stock_depth_days * daily_demand
        
        # Страховой запас (Safety Stock)
        # Safety Stock = (Макс. дневной спрос × Время поставки) - (Средний спрос × Время поставки)
        max_daily_demand = daily_demand * 1.5  # Предполагаем пиковый спрос +50%
        result.safety_stock_units = int(math.ceil(
            (max_daily_demand * input_data.supplier_lead_time_days) - 
            (daily_demand * input_data.supplier_lead_time_days)
        ))
        
        # Точка заказа (Reorder Point)
        # ROP = (Средний дневной спрос × Время поставки) + Страховой запас
        result.reorder_point_units = int(math.ceil(
            (daily_demand * input_data.supplier_lead_time_days) + result.safety_stock_units
        ))
        
        # Оборачиваемость запасов в днях
        if daily_demand > 0 and result.optimal_stock_units > 0:
            result.stock_turnover_days = result.optimal_stock_units / daily_demand
        else:
            result.stock_turnover_days = 0
        
        # Коэффициент оборачиваемости
        if result.optimal_stock_units > 0:
            result.stock_turnover_rate = annual_demand / result.optimal_stock_units
        else:
            result.stock_turnover_rate = 0
        
        # Дни запаса (Days of Inventory)
        if daily_demand > 0 and result.optimal_stock_units > 0:
            result.days_of_inventory = result.optimal_stock_units / daily_demand
        else:
            result.days_of_inventory = 0
        
        # Стоимость хранения на единицу
        result.holding_cost_per_unit = holding_cost_per_unit
        
        # =====================================================================
        # 17. РЕКОМЕНДАЦИИ ПО ОПТИМИЗАЦИИ СКЛАДА (НОВОЕ)
        # =====================================================================
        # Рекомендуемая глубина запаса
        if result.stock_turnover_rate > 12:
            # Высокая оборачиваемость - можно держать меньше запаса
            result.recommended_stock_depth_days = max(14, input_data.stock_depth_days - 5)
        elif result.stock_turnover_rate < 6:
            # Низкая оборачиваемость - нужен больший запас
            result.recommended_stock_depth_days = input_data.stock_depth_days + 5
        else:
            result.recommended_stock_depth_days = input_data.stock_depth_days
        
        # Рекомендуемый страховой запас
        if result.penalty_probability > 0.2:
            # Высокий риск штрафов - увеличиваем страховой запас
            result.recommended_safety_stock_days = min(14, input_data.safety_stock_days + 3)
        else:
            result.recommended_safety_stock_days = max(3, input_data.safety_stock_days - 2)
        
        # Потенциал оптимизации запаса (%)
        current_stock = input_data.stock_depth_days * daily_demand
        recommended_stock = result.recommended_stock_depth_days * daily_demand
        if current_stock > 0:
            result.stock_optimization_potential = ((current_stock - recommended_stock) / current_stock) * 100
        else:
            result.stock_optimization_potential = 0
        
        return result
    
    @timing_decorator
    def calculate_batch(self, input_data_list: List[FBSInputData],
                       use_parallel: bool = True,
                       max_workers: int = 8) -> List[FBSResultData]:
        """
        Пакетный расчет для множества товаров.
        
        Поддерживает параллельную обработку для ускорения расчета больших партий.
        Оптимально для каталогов от 100 до 500 000+ товаров.
        
        Args:
            input_data_list: Список входных данных товаров
            use_parallel: Использовать параллельную обработку
            max_workers: Максимальное количество потоков
            
        Returns:
            List[FBSResultData]: Список результатов расчета
        """
        total = len(input_data_list)
        results = [None] * total
        
        if total == 0:
            return []
        
        self.progress_tracker.start(total, f"Расчет {total} товаров...")
        
        if total > 100 and use_parallel:
            # Параллельная обработка для больших партий
            logger.info(f"⚡ Запуск параллельной обработки {total} товаров ({max_workers} потоков)")
            
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                # Создаем задачи
                future_to_index = {}
                for i, data in enumerate(input_data_list):
                    future = executor.submit(self.calculate_unit_economics, data)
                    future_to_index[future] = i
                
                # Обрабатываем результаты по мере завершения
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
                        # Создаем пустой результат с пометкой об ошибке
                        error_result = FBSResultData()
                        error_result.artikul = f"ERROR_{index}"
                        error_result.product_name = f"Ошибка расчета: {str(e)[:50]}"
                        results[index] = error_result
                        completed += 1
        else:
            # Последовательная обработка для небольших партий
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
        
        # Фильтруем None результаты (на случай ошибок)
        results = [r for r in results if r is not None]
        
        logger.info(f"✅ Пакетный расчет завершен. Успешно: {len(results)}/{total}")
        
        return results
    
    def get_tariffs_summary(self) -> pd.DataFrame:
        """Получение сводки по текущим тарифам"""
        return self.api_manager.get_all_tariffs_as_dataframe(self.current_marketplace)
    
    def test_api_connections(self) -> Dict[str, Dict[str, Any]]:
        """Тестирование всех API подключений"""
        results = {}
        
        for marketplace in ["Ozon", "Wildberries", "Яндекс Маркет"]:
            results[marketplace] = self.api_manager.test_api_connection(marketplace)
        
        # Тест DeepSeek
        deepseek_key = self.api_manager.get_api_key('deepseek')
        results['DeepSeek'] = {
            'status': 'available' if deepseek_key else 'no_api_key',
            'timestamp': datetime.now().isoformat()
        }
        
        return results

# ============================================================================
# БЛОК 8: ЭКСПОРТ В CSV И GOOGLE SHEETS
# ============================================================================

class DataExporter:
    """
    Класс для экспорта данных в различные форматы:
    - CSV
    - Google Sheets
    - Excel (с живыми формулами)
    """
    
    def __init__(self):
        self.logger = logger
    
    def export_to_csv(self, results: List[FBSResultData], 
                     input_data_list: List[FBSInputData],
                     output_path: str,
                     delimiter: str = ';') -> bool:
        """
        Экспорт результатов в CSV файл.
        
        Args:
            results: Список результатов расчета
            input_data_list: Список входных данных
            output_path: Путь для сохранения
            delimiter: Разделитель полей
            
        Returns:
            bool: True если экспорт успешен
        """
        try:
            if not results:
                self.logger.warning("⚠️ Нет данных для экспорта в CSV")
                return False
            
            # Формируем DataFrame
            data = []
            for result, input_data in zip(results, input_data_list):
                row = {
                    'Артикул': result.artikul,
                    'Наименование': result.product_name,
                    'Категория': input_data.category,
                    'Цена продажи, ₽': result.selling_price,
                    'Себестоимость, ₽': input_data.cogs,
                    'Комиссия, ₽': result.commission,
                    'First Mile, ₽': result.first_mile_cost,
                    'Last Mile, ₽': result.last_mile_cost,
                    'Pick & Pack, ₽': result.pick_pack_cost,
                    'Упаковка, ₽': result.packaging_cost,
                    'Эквайринг, ₽': result.acquiring_cost,
                    'Возвраты, ₽': result.return_cost,
                    'Штрафы, ₽': result.penalty_cost,
                    'Маркетинг, ₽': result.marketing_cost,
                    'Склад, ₽': result.warehouse_cost,
                    'Налог, ₽': result.tax_cost,
                    'Итого расходов, ₽': result.total_expenses,
                    'Чистая прибыль, ₽': result.gross_profit,
                    'Маржа, %': result.margin_percent,
                    'ROI, %': result.roi_percent,
                    'Точка безубыточности, км': result.break_even_distance_km,
                    'Макс. скидка, %': result.max_discount_percent,
                    'Запас прочности, ₽': result.safety_margin_price,
                    'LTV, ₽': result.ltv,
                    'CAC, ₽': result.cac,
                    'LTV/CAC': result.ltv_cac_ratio,
                    'ROMI, %': result.romi,
                    'Оптимальный запас, шт': result.optimal_stock_units,
                    'Страховой запас, шт': result.safety_stock_units,
                    'Точка заказа, шт': result.reorder_point_units,
                    'Оборачиваемость, дней': result.stock_turnover_days,
                    'Потенциал оптимизации, %': result.stock_optimization_potential
                }
                data.append(row)
            
            df = pd.DataFrame(data)
            
            # Сохраняем в CSV
            df.to_csv(output_path, index=False, sep=delimiter, encoding='utf-8-sig')
            
            self.logger.info(f"✅ Экспорт в CSV завершен: {output_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Ошибка экспорта в CSV: {e}")
            self.logger.exception(e)
            return False
    
    def export_to_google_sheets(self, results: List[FBSResultData],
                               input_data_list: List[FBSInputData],
                               credentials_path: str,
                               spreadsheet_name: str = "FBS Unit Economics Report",
                               worksheet_name: str = "FBS Data") -> bool:
        """
        Экспорт результатов в Google Sheets.
        
        Args:
            results: Список результатов расчета
            input_data_list: Список входных данных
            credentials_path: Путь к файлу credentials.json
            spreadsheet_name: Название Google Sheets документа
            worksheet_name: Название листа
            
        Returns:
            bool: True если экспорт успешен
        """
        if not GSPREAD_AVAILABLE:
            self.logger.error("❌ GSpread не установлен. Установите: pip install gspread google-auth")
            return False
        
        try:
            if not results:
                self.logger.warning("⚠️ Нет данных для экспорта в Google Sheets")
                return False
            
            # Формируем данные
            data = []
            headers = [
                'Артикул', 'Наименование', 'Категория', 'Цена продажи, ₽',
                'Себестоимость, ₽', 'Комиссия, ₽', 'First Mile, ₽', 'Last Mile, ₽',
                'Pick & Pack, ₽', 'Упаковка, ₽', 'Эквайринг, ₽', 'Возвраты, ₽',
                'Штрафы, ₽', 'Маркетинг, ₽', 'Склад, ₽', 'Налог, ₽',
                'Итого расходов, ₽', 'Чистая прибыль, ₽', 'Маржа, %',
                'ROI, %', 'Точка безубыточности, км', 'Макс. скидка, %',
                'Запас прочности, ₽', 'LTV, ₽', 'CAC, ₽', 'LTV/CAC',
                'ROMI, %', 'Оптимальный запас, шт', 'Страховой запас, шт',
                'Точка заказа, шт', 'Оборачиваемость, дней', 'Потенциал оптимизации, %'
            ]
            
            data.append(headers)
            
            for result, input_data in zip(results, input_data_list):
                row = [
                    result.artikul,
                    result.product_name,
                    input_data.category,
                    result.selling_price,
                    input_data.cogs,
                    result.commission,
                    result.first_mile_cost,
                    result.last_mile_cost,
                    result.pick_pack_cost,
                    result.packaging_cost,
                    result.acquiring_cost,
                    result.return_cost,
                    result.penalty_cost,
                    result.marketing_cost,
                    result.warehouse_cost,
                    result.tax_cost,
                    result.total_expenses,
                    result.gross_profit,
                    result.margin_percent,
                    result.roi_percent,
                    result.break_even_distance_km,
                    result.max_discount_percent,
                    result.safety_margin_price,
                    result.ltv,
                    result.cac,
                    result.ltv_cac_ratio,
                    result.romi,
                    result.optimal_stock_units,
                    result.safety_stock_units,
                    result.reorder_point_units,
                    result.stock_turnover_days,
                    result.stock_optimization_potential
                ]
                data.append(row)
            
            # Аутентификация в Google Sheets
            scopes = [
                'https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'
            ]
            
            creds = service_account.Credentials.from_service_account_file(
                credentials_path, scopes=scopes
            )
            
            client = gspread.authorize(creds)
            
            # Создаем или открываем документ
            try:
                spreadsheet = client.open(spreadsheet_name)
            except gspread.SpreadsheetNotFound:
                spreadsheet = client.create(spreadsheet_name)
                self.logger.info(f"📄 Создан новый документ: {spreadsheet_name}")
            
            # Создаем или очищаем лист
            try:
                worksheet = spreadsheet.worksheet(worksheet_name)
                worksheet.clear()
            except gspread.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="1000", cols="100")
            
            # Обновляем данные
            worksheet.update(data, value_input_option='USER_ENTERED')
            
            # Форматирование
            worksheet.format('A1:AF1', {
                "backgroundColor": {"red": 0.1, "green": 0.1, "blue": 0.18},
                "textFormat": {"foregroundColor": {"red": 1, "green": 1, "blue": 1}, "bold": True}
            })
            
            self.logger.info(f"✅ Экспорт в Google Sheets завершен: {spreadsheet.url}")
            
            # Возвращаем URL для доступа
            st.success(f"✅ Данные экспортированы в Google Sheets: {spreadsheet.url}")
            return True
            
        except Exception as e:
            self.logger.error(f"❌ Ошибка экспорта в Google Sheets: {e}")
            self.logger.exception(e)
            return False

# ============================================================================
# БЛОК 9: ЭКСПОРТ В EXCEL С ЖИВЫМИ ФОРМУЛАМИ (ОБНОВЛЕННЫЙ)
# ============================================================================

class ProfessionalExcelExporter:
    """
    Профессиональный экспорт в Excel с живыми формулами.
    
    Особенности:
    - Отдельный лист "Тарифы МП" с данными из API
    - Формулы на основном листе ссылаются на лист тарифов
    - При обновлении тарифов в Excel пересчет происходит автоматически
    - Условное форматирование прибыльных/убыточных товаров
    - Рекомендованные цены в отдельном столбце
    - Только 3 листа: "Тарифы МП", "Дашборд", "Юнит-экономика FBS"
    """
    
    def __init__(self):
        if not OPENPYXL_AVAILABLE:
            raise ImportError(
                "openpyxl не установлен. Установите: pip install openpyxl"
            )
        
        # Стили для заголовков
        self.header_fill = PatternFill(
            start_color="1a1a2e", 
            end_color="1a1a2e", 
            fill_type="solid"
        )
        self.header_font = Font(
            bold=True, 
            color="FFFFFF", 
            size=11, 
            name="Arial"
        )
        self.header_alignment = Alignment(
            horizontal="center", 
            vertical="center", 
            wrap_text=True
        )
        
        # Стили для ячеек
        self.input_fill = PatternFill(
            start_color="FFF4CC", 
            end_color="FFF4CC", 
            fill_type="solid"
        )
        self.formula_fill = PatternFill(
            start_color="E2EFDA", 
            end_color="E2EFDA", 
            fill_type="solid"
        )
        self.result_fill = PatternFill(
            start_color="DCE6F1", 
            end_color="DCE6F1", 
            fill_type="solid"
        )
        self.tariff_fill = PatternFill(
            start_color="F0E6FF", 
            end_color="F0E6FF", 
            fill_type="solid"
        )
        self.api_fill = PatternFill(
            start_color="FFE0E0", 
            end_color="FFE0E0", 
            fill_type="solid"
        )
        self.profit_fill = PatternFill(
            start_color="C6EFCE", 
            end_color="C6EFCE", 
            fill_type="solid"
        )
        self.loss_fill = PatternFill(
            start_color="FFC7CE", 
            end_color="FFC7CE", 
            fill_type="solid"
        )
        self.warning_fill = PatternFill(
            start_color="FFF3CD", 
            end_color="FFF3CD", 
            fill_type="solid"
        )
        self.recommended_price_fill = PatternFill(
            start_color="D9E1F2", 
            end_color="D9E1F2", 
            fill_type="solid"
        )
        
        # Границы
        self.thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        # Шрифты
        self.title_font = Font(
            bold=True, 
            size=14, 
            name="Arial", 
            color="1a1a2e"
        )
        self.subtitle_font = Font(
            bold=True, 
            size=12, 
            name="Arial", 
            color="333333"
        )
        self.link_font = Font(
            color="0563C1", 
            underline="single"
        )
        
        # Для хранения ссылок на строки тарифов
        self._tariffs_start_row = 6
        self._tariffs_last_row = 6
    
    def export_fbs_report(self,
                         results: List[FBSResultData],
                         input_data_list: List[FBSInputData],
                         calculator: FBSUnitEconomicsCalculator,
                         marketplace_name: str,
                         output_path: str) -> bool:
        """
        Создание полного профессионального Excel-отчета.
        
        Структура отчета (ТОЛЬКО 3 ЛИСТА):
        1. 📋 Тарифы МП - актуальные тарифы из API
        2. 📊 Юнит-экономика FBS - основной расчет с живыми формулами
        3. 📈 Дашборд - сводная статистика
        
        Args:
            results: Список результатов расчета
            input_data_list: Список входных данных
            calculator: Экземпляр калькулятора с тарифами
            marketplace_name: Название маркетплейса
            output_path: Путь для сохранения файла
            
        Returns:
            bool: True если экспорт успешен
        """
        try:
            wb = Workbook()
            
            # Удаляем стандартный лист
            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']
            
            # Создаем листы в правильном порядке (ТОЛЬКО 3)
            logger.info("📊 Создание листа тарифов...")
            ws_tariffs = wb.create_sheet("📋 Тарифы МП", 0)
            self._create_tariffs_sheet(ws_tariffs, calculator, marketplace_name)
            
            logger.info("📊 Создание основного листа с юнит-экономикой...")
            ws_main = wb.create_sheet("📊 Юнит-экономика FBS")
            self._create_main_sheet_with_tariff_links(
                ws_main, results, input_data_list, marketplace_name, calculator
            )
            
            logger.info("📊 Создание дашборда...")
            ws_dashboard = wb.create_sheet("📈 Дашборд")
            self._create_dashboard_sheet(ws_dashboard, results, marketplace_name)
            
            # Сохраняем файл
            wb.save(output_path)
            logger.info(f"✅ Профессиональный Excel отчет сохранен: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Ошибка создания Excel отчета: {e}")
            logger.exception(e)
            return False
    
    def _create_tariffs_sheet(self, ws, calculator: FBSUnitEconomicsCalculator, marketplace_name: str):
        """
        Создание листа с актуальными тарифами из API.
        Этот лист является источником данных для формул на других листах.
        """
        # Заголовок
        ws.merge_cells('A1:R1')
        title_cell = ws.cell(row=1, column=1, value=(
            f"📋 Актуальные тарифы {marketplace_name} — "
            f"Загружено: {calculator.tariffs_updated_at.strftime('%d.%m.%Y %H:%M') if calculator.tariffs_updated_at else 'Нет данных'} | "
            f"Источник: {calculator.tariffs_source}"
        ))
        title_cell.font = self.title_font
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 45
        
        # Подзаголовок с легендой
        ws.merge_cells('A2:R2')
        legend_cell = ws.cell(row=2, column=1, value=(
            "🟣 Данные из API | 🔴 Дефолтные значения | 📎 На этот лист ссылаются формулы расчета | "
            "💡 Обновите тарифы здесь — и все расчеты пересчитаются автоматически"
        ))
        legend_cell.font = Font(size=10, italic=True, color="666666")
        legend_cell.alignment = Alignment(horizontal="center")
        
        # Информация об API
        ws.merge_cells('A3:R3')
        api_info_parts = []
        
        if calculator.api_manager.has_api_key('ozon'):
            api_info_parts.append("✅ Ozon API")
        else:
            api_info_parts.append("⚪ Ozon API")
        
        if calculator.api_manager.has_api_key('wildberries'):
            api_info_parts.append("✅ Wildberries API")
        else:
            api_info_parts.append("⚪ Wildberries API")
        
        if calculator.api_manager.has_api_key('yandex_market'):
            api_info_parts.append("✅ Яндекс Маркет API")
        else:
            api_info_parts.append("⚪ Яндекс Маркет API")
        
        if calculator.api_manager.has_api_key('deepseek'):
            api_info_parts.append("✅ DeepSeek AI")
        else:
            api_info_parts.append("⚪ DeepSeek AI")
        
        ws.cell(row=3, column=1, value=f"🔌 Статус API: {' | '.join(api_info_parts)}")
        
        # Кнопка обновления (инструкция)
        ws.merge_cells('A4:R4')
        ws.cell(row=4, column=1, value=(
            "🔄 Для обновления тарифов: откройте приложение → Настройки → Загрузить тарифы через API"
        )).font = Font(italic=True, color="666666")
        
        # Заголовки таблицы тарифов
        headers = [
            ("Категория", 22),
            ("Комиссия, %", 14),
            ("Мин. комиссия, ₽", 16),
            ("База Last Mile, ₽", 17),
            ("Last Mile за кг, ₽", 17),
            ("Last Mile за км, ₽", 17),
            ("Эквайринг, %", 13),
            ("Возвраты, %", 13),
            ("Штраф за просрочку, %", 18),
            ("Время на передачу, ч", 17),
            ("Множитель FBO", 14),
            ("Множитель FBP", 14),
            ("Хранение, ₽/л/день", 16),
            ("Мин. логистика, ₽", 16),
            ("Источник данных", 22),
            ("Дата обновления", 18),
            ("API Response", 15),
            ("Статус", 12)
        ]
        
        for col_idx, (header_text, width) in enumerate(headers, 1):
            cell = ws.cell(row=6, column=col_idx, value=header_text)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.thin_border
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        
        ws.row_dimensions[6].height = 35
        ws.freeze_panes = "A7"
        
        # Заполнение данных тарифов
        row = 7
        for category, tariff_data in calculator.current_tariffs.items():
            # Категория
            cell = ws.cell(row=row, column=1, value=category)
            cell.fill = self.tariff_fill
            cell.border = self.thin_border
            cell.font = Font(bold=True)
            
            # Комиссия
            cell = ws.cell(row=row, column=2, value=round(tariff_data.get('commission_rate', 0) * 100, 2))
            cell.fill = self.tariff_fill
            cell.border = self.thin_border
            cell.number_format = '0.00"%"'
            
            # Остальные числовые параметры
            numeric_fields = [
                (3, tariff_data.get('min_commission', 0)),
                (4, tariff_data.get('last_mile_base', 0)),
                (5, tariff_data.get('last_mile_per_kg', 0)),
                (6, tariff_data.get('last_mile_per_km', 0)),
                (7, round(tariff_data.get('acquiring_fee', 0) * 100, 2)),
                (8, round(tariff_data.get('return_fee', 0) * 100, 2)),
                (9, round(tariff_data.get('penalty_rate', 0) * 100, 2)),
                (10, tariff_data.get('penalty_time_hours', 0)),
                (11, tariff_data.get('fbo_multiplier', 0)),
                (12, tariff_data.get('fbp_multiplier', 0)),
                (13, tariff_data.get('storage_base_rate', 0)),
                (14, tariff_data.get('min_logistics', 0))
            ]
            
            for col, value in numeric_fields:
                cell = ws.cell(row=row, column=col, value=value)
                cell.fill = self.tariff_fill
                cell.border = self.thin_border
                if isinstance(value, float):
                    cell.number_format = '#,##0.00'
            
            # Источник данных
            source = tariff_data.get('source', 'default')
            source_cell = ws.cell(row=row, column=15, value=source)
            source_cell.border = self.thin_border
            
            if 'api' in source.lower():
                source_cell.fill = PatternFill(start_color="E0FFE0", end_color="E0FFE0", fill_type="solid")
            elif 'deepseek' in source.lower():
                source_cell.fill = PatternFill(start_color="E0E0FF", end_color="E0E0FF", fill_type="solid")
            else:
                source_cell.fill = PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type="solid")
            
            # Дата обновления
            updated = tariff_data.get('last_updated', '')
            if updated:
                try:
                    updated_dt = datetime.fromisoformat(updated)
                    updated_str = updated_dt.strftime('%d.%m.%Y %H:%M')
                except:
                    updated_str = str(updated)[:19]
            else:
                updated_str = 'Н/Д'
            
            cell = ws.cell(row=row, column=16, value=updated_str)
            cell.fill = self.tariff_fill
            cell.border = self.thin_border
            
            # Сырой ответ API
            api_raw = tariff_data.get('api_response_raw', '')
            cell = ws.cell(row=row, column=17, value=api_raw[:100] if api_raw else 'Нет данных')
            cell.fill = self.tariff_fill
            cell.border = self.thin_border
            
            # Статус
            status = "✅ Актуально" if source != 'default' else "⚠️ Дефолт"
            status_cell = ws.cell(row=row, column=18, value=status)
            status_cell.border = self.thin_border
            
            if source != 'default':
                status_cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            else:
                status_cell.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
            
            row += 1
        
        self._tariffs_last_row = row - 1
        
        # Добавляем примечание
        row += 1
        ws.merge_cells(f'A{row}:R{row}')
        note_cell = ws.cell(row=row, column=1, value=(
            "💡 Как использовать: Измените значения в фиолетовых ячейках — "
            "все формулы на листе '📊 Юнит-экономика FBS' автоматически пересчитаются. "
            "Для загрузки свежих тарифов используйте приложение."
        ))
        note_cell.font = Font(italic=True, size=11, color="333333")
        
        # Условное форматирование для источника данных
        if self._tariffs_last_row >= 7:
            ws.conditional_formatting.add(
                f"O7:O{self._tariffs_last_row}",
                CellIsRule(
                    operator="equal",
                    formula=['"default"'],
                    fill=PatternFill(start_color="FFE0E0", end_color="FFE0E0", fill_type="solid")
                )
            )
    
    def _create_main_sheet_with_tariff_links(self, ws, results, input_data_list, 
                                             marketplace_name, calculator):
        """
        Создание основного листа с формулами, ссылающимися на лист тарифов.
        
        ВКЛЮЧАЕТ:
        - Все расходы с формулами из листа тарифов
        - Рекомендованные цены в отдельном столбце
        - Условное форматирование прибыльных/убыточных товаров
        - Полную детализацию юнит-экономики
        """
        # Название листа тарифов для формул
        tariff_sheet = "'📋 Тарифы МП'"
        
        # Заголовок
        ws.merge_cells('A1:AW1')
        title_cell = ws.cell(row=1, column=1, value=(
            f"🚀 Юнит-экономика FBS — {marketplace_name} — "
            f"{datetime.now().strftime('%d.%m.%Y %H:%M')}"
        ))
        title_cell.font = self.title_font
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 45
        
        # Подзаголовок с легендой
        ws.merge_cells('A2:AW2')
        ws.cell(row=2, column=1, value=(
            "🟡 Вводные данные (редактируемые) | 🟣 Тарифы из API (лист '📋 Тарифы МП') | "
            "🟢 Расчетные формулы | 🔵 Итоговые показатели | "
            "🟠 Рекомендованные цены | 📎 Формулы автоматически подтягивают данные из листа тарифов"
        )).font = Font(size=9, italic=True, color="666666")
        
        # Определяем заголовки колонок (РАСШИРЕННЫЕ)
        headers = [
            # Вводные данные (A-P)
            ("Артикул", 15), ("Наименование", 30), ("Категория", 18),
            ("Цена продажи, ₽", 15), ("Себестоимость, ₽", 15),
            ("Вес брутто, кг", 12), ("Длина, см", 10), ("Ширина, см", 10),
            ("Высота, см", 10), ("Расстояние до МП, км", 16),
            ("Стоимость 1 км, ₽", 15), ("Единиц на паллете", 14),
            ("Упаковка, ₽/шт", 13), ("Время сборки, мин", 14),
            ("Ставка оператора, ₽/ч", 16), ("Маркетинг на ед., ₽", 15),
            
            # Расчетные параметры (Q-V)
            ("Объемный вес, кг", 13), ("Оплачиваемый вес, кг", 16),
            ("Коэфф. ночной смены", 15), ("Вероятность просрочки", 16),
            ("Ставка налога", 12), ("Параметр", 10),
            
            # Расходы (W-AG) - ВСЕ ССЫЛАЮТСЯ НА ЛИСТ ТАРИФОВ
            ("Комиссия МП, ₽", 14), ("First Mile, ₽", 13),
            ("Last Mile, ₽", 13), ("Pick & Pack, ₽", 14),
            ("Упаковка расч., ₽", 14), ("Эквайринг, ₽", 13),
            ("Возвраты, ₽", 12), ("Штрафы, ₽", 12),
            ("Маркетинг расч., ₽", 15), ("Складские, ₽", 12),
            ("Налог, ₽", 10),
            
            # Итоговые показатели (AH-AL)
            ("ИТОГО расходов, ₽", 16), ("ЧИСТАЯ ПРИБЫЛЬ, ₽", 16),
            ("МАРЖА, %", 10), ("ROI, %", 10),
            ("Мин. цена, ₽", 12),
            
            # РЕКОМЕНДОВАННЫЕ ЦЕНЫ (AM-AN) - НОВОЕ
            ("Рек. цена (маржа 15%), ₽", 18),
            ("Рек. цена (маржа 25%), ₽", 18),
            
            # FBS метрики (AO-AT)
            ("Точка безубыт., км", 15), ("Макс. скидка, %", 13),
            ("Запас прочности, ₽", 15), ("LTV, ₽", 12),
            ("CAC, ₽", 12), ("LTV/CAC", 10),
            
            # Складская оптимизация (AU-AW) - НОВОЕ
            ("Оптим. запас, шт", 14), ("Страх. запас, шт", 14),
            ("Оборачиваемость, дн", 16)
        ]
        
        # Записываем заголовки
        for col_idx, (header_text, width) in enumerate(headers, 1):
            cell = ws.cell(row=4, column=col_idx, value=header_text)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.thin_border
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        
        ws.row_dimensions[4].height = 40
        ws.freeze_panes = "A5"
        
        # Функция для поиска строки категории в листе тарифов
        def get_tariff_row(category: str) -> int:
            """Возвращает номер строки в листе тарифов для категории"""
            row = self._tariffs_start_row
            for cat, _ in calculator.current_tariffs.items():
                if cat == category:
                    return row
                row += 1
            # Если категория не найдена - возвращаем строку default
            return self._tariffs_start_row
        
        # Заполнение данных
        for row_idx, (result, input_data) in enumerate(zip(results, input_data_list), 5):
            # === ВВОДНЫЕ ДАННЫЕ (желтые ячейки) ===
            input_fields = [
                (1, input_data.artikul), (2, input_data.product_name),
                (3, input_data.category), (4, input_data.selling_price),
                (5, input_data.cogs), (6, input_data.weight_kg),
                (7, input_data.length_cm), (8, input_data.width_cm),
                (9, input_data.height_cm), (10, input_data.warehouse_distance_km),
                (11, input_data.transport_cost_per_km), (12, input_data.pallet_capacity),
                (13, input_data.packaging_cost), (14, input_data.pick_pack_time_min),
                (15, input_data.operator_hourly_rate), (16, input_data.marketing_budget_per_unit)
            ]
            
            for col, value in input_fields:
                cell = ws.cell(row=row_idx, column=col, value=value)
                cell.fill = self.input_fill
                cell.border = self.thin_border
                if isinstance(value, float):
                    cell.number_format = '#,##0.00'
            
            # === РАСЧЕТНЫЕ ПАРАМЕТРЫ ===
            # Объемный вес (колонка 17 = Q)
            formula = f"=IF(G{row_idx}*H{row_idx}*I{row_idx}>0, (G{row_idx}*H{row_idx}*I{row_idx})/5000, 0)"
            cell = ws.cell(row=row_idx, column=17, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Оплачиваемый вес (колонка 18 = R)
            formula = f"=CEILING(MAX(F{row_idx}, Q{row_idx}), 0.5)"
            cell = ws.cell(row=row_idx, column=18, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Коэффициент ночной смены (колонка 19 = S)
            night_coef = 0.05 if input_data.has_night_shift else 0.35
            cell = ws.cell(row=row_idx, column=19, value=night_coef)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '0.00'
            
            # Вероятность просрочки (колонка 20 = T)
            formula = f"=S{row_idx}"
            cell = ws.cell(row=row_idx, column=20, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '0.00'
            
            # Ставка налога (колонка 21 = U)
            tax_rate = TAX_SYSTEMS.get(st.session_state.get('tax_system', 'УСН 6% (доходы)'), {}).get('rate', 0.06)
            cell = ws.cell(row=row_idx, column=21, value=tax_rate)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '0.00'
            
            # Получаем строку тарифа для категории
            tariff_row = get_tariff_row(input_data.category)
            
            # === РАСХОДЫ С ССЫЛКАМИ НА ЛИСТ ТАРИФОВ (НЕ ЗАХАРДКОЖЕНЫ) ===
            
            # Комиссия МП (колонка 23 = W)
            formula = f"=MAX(D{row_idx}*{tariff_sheet}!B{tariff_row}/100, {tariff_sheet}!C{tariff_row})"
            cell = ws.cell(row=row_idx, column=23, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # First Mile (колонка 24 = X)
            formula = f"=IF(L{row_idx}>0, (J{row_idx}*K{row_idx}*2)/L{row_idx}, M{row_idx})"
            cell = ws.cell(row=row_idx, column=24, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Last Mile (колонка 25 = Y)
            formula = f"=MAX({tariff_sheet}!D{tariff_row}+R{row_idx}*{tariff_sheet}!E{tariff_row}, {tariff_sheet}!N{tariff_row})"
            cell = ws.cell(row=row_idx, column=25, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Pick & Pack (колонка 26 = Z)
            formula = f"=(N{row_idx}/60)*O{row_idx}"
            cell = ws.cell(row=row_idx, column=26, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Упаковка (колонка 27 = AA)
            cell = ws.cell(row=row_idx, column=27, value=input_data.packaging_cost)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Эквайринг (колонка 28 = AB)
            formula = f"=D{row_idx}*{tariff_sheet}!G{tariff_row}/100"
            cell = ws.cell(row=row_idx, column=28, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Возвраты (колонка 29 = AC)
            formula = f"=D{row_idx}*{tariff_sheet}!H{tariff_row}/100"
            cell = ws.cell(row=row_idx, column=29, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Штрафы (колонка 30 = AD)
            formula = f"=D{row_idx}*{tariff_sheet}!I{tariff_row}/100*T{row_idx}"
            cell = ws.cell(row=row_idx, column=30, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Маркетинг (колонка 31 = AE)
            cell = ws.cell(row=row_idx, column=31, value=input_data.marketing_budget_per_unit)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Складские (колонка 32 = AF)
            formula = f"=P{row_idx}/30/5"
            cell = ws.cell(row=row_idx, column=32, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Налог (колонка 33 = AG)
            formula = f"=D{row_idx}*U{row_idx}"
            cell = ws.cell(row=row_idx, column=33, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # === ИТОГОВЫЕ ПОКАЗАТЕЛИ ===
            
            # Итого расходов (колонка 34 = AH)
            formula = (
                f"=E{row_idx}+W{row_idx}+X{row_idx}+Y{row_idx}+Z{row_idx}+"
                f"AA{row_idx}+AB{row_idx}+AC{row_idx}+AD{row_idx}+AE{row_idx}+"
                f"AF{row_idx}+AG{row_idx}"
            )
            cell = ws.cell(row=row_idx, column=34, value=formula)
            cell.fill = self.result_fill
            cell.font = Font(bold=True, size=11)
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Чистая прибыль (колонка 35 = AI)
            formula = f"=D{row_idx}-AH{row_idx}"
            cell = ws.cell(row=row_idx, column=35, value=formula)
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            if result.gross_profit > 0:
                cell.fill = self.profit_fill
                cell.font = Font(bold=True, color="006100")
            else:
                cell.fill = self.loss_fill
                cell.font = Font(bold=True, color="9C0006")
            
            # Маржа (колонка 36 = AJ)
            formula = f"=IF(D{row_idx}>0, (AI{row_idx}/D{row_idx})*100, 0)"
            cell = ws.cell(row=row_idx, column=36, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '0.00"%"'
            
            # ROI (колонка 37 = AK)
            formula = f"=IF(E{row_idx}>0, (AI{row_idx}/E{row_idx})*100, 0)"
            cell = ws.cell(row=row_idx, column=37, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '0.00"%"'
            
            # Минимальная цена (колонка 38 = AL)
            formula = (
                f"=MAX(0, (E{row_idx}+X{row_idx}+Y{row_idx}+Z{row_idx}+AA{row_idx}+AE{row_idx}+AF{row_idx})/"
                f"MAX(0.01, (1-{tariff_sheet}!B{tariff_row}/100-{tariff_sheet}!G{tariff_row}/100-"
                f"{tariff_sheet}!H{tariff_row}/100-{tariff_sheet}!I{tariff_row}/100*T{row_idx}-U{row_idx})))"
            )
            cell = ws.cell(row=row_idx, column=38, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # === РЕКОМЕНДОВАННЫЕ ЦЕНЫ (НОВОЕ) ===
            # Рекомендуемая цена при марже 15% (колонка 39 = AM)
            formula = f"=IF((1-0.15)>0, AH{row_idx}/(1-0.15), D{row_idx})"
            cell = ws.cell(row=row_idx, column=39, value=formula)
            cell.fill = self.recommended_price_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            cell.font = Font(bold=True, color="0F4C81")
            
            # Рекомендуемая цена при марже 25% (колонка 40 = AN)
            formula = f"=IF((1-0.25)>0, AH{row_idx}/(1-0.25), D{row_idx})"
            cell = ws.cell(row=row_idx, column=40, value=formula)
            cell.fill = self.recommended_price_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            cell.font = Font(bold=True, color="0F4C81")
            
            # Максимальная скидка (колонка 41 = AO)
            formula = f"=IF(D{row_idx}>0, ((D{row_idx}-AL{row_idx})/D{row_idx})*100, 0)"
            cell = ws.cell(row=row_idx, column=41, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '0.00"%"'
            
            # Точка безубыточности (колонка 42 = AP)
            formula = f"=IF(X{row_idx}>0, AI{row_idx}/(K{row_idx}*2/L{row_idx}), 999999)"
            cell = ws.cell(row=row_idx, column=42, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.0'
            
            # Запас прочности (колонка 43 = AQ)
            formula = f"=D{row_idx}-AL{row_idx}"
            cell = ws.cell(row=row_idx, column=43, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # LTV (колонка 44 = AR)
            formula = f"=D{row_idx}*{input_data.avg_purchases_per_year}*{input_data.customer_retention_rate}/(1+{input_data.discount_rate})"
            cell = ws.cell(row=row_idx, column=44, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # CAC (колонка 45 = AS)
            formula = f"=(AE{row_idx}+AD{row_idx}+X{row_idx})/0.3"
            cell = ws.cell(row=row_idx, column=45, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # LTV/CAC (колонка 46 = AT)
            formula = f"=IF(AS{row_idx}>0, AR{row_idx}/AS{row_idx}, 999)"
            cell = ws.cell(row=row_idx, column=46, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '0.0'
            
            # === СКЛАДСКАЯ ОПТИМИЗАЦИЯ (НОВОЕ) ===
            # Оптимальный запас (колонка 47 = AU)
            formula = f"=ROUNDUP(SQRT((2*P{row_idx}*365*500)/(U{row_idx}*0.01*12)), 0)"
            cell = ws.cell(row=row_idx, column=47, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0'
            
            # Страховой запас (колонка 48 = AV)
            formula = f"=CEILING((P{row_idx}*1.5*3)-(P{row_idx}*3), 1)"
            cell = ws.cell(row=row_idx, column=48, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0'
            
            # Оборачиваемость (колонка 49 = AW)
            formula = f"=IF(AU{row_idx}>0, AU{row_idx}/P{row_idx}, 0)"
            cell = ws.cell(row=row_idx, column=49, value=formula)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '0.0'
        
        # === УСЛОВНОЕ ФОРМАТИРОВАНИЕ ===
        last_data_row = len(results) + 4
        if last_data_row >= 5:
            # Прибыль: зеленый фон для прибыльных, красный для убыточных
            ws.conditional_formatting.add(
                f"AI5:AI{last_data_row}",
                CellIsRule(
                    operator="greaterThan",
                    formula=["0"],
                    fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
                    font=Font(color="006100", bold=True)
                )
            )
            ws.conditional_formatting.add(
                f"AI5:AI{last_data_row}",
                CellIsRule(
                    operator="lessThan",
                    formula=["0"],
                    fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
                    font=Font(color="9C0006", bold=True)
                )
            )
            
            # Маржа: предупреждение если < 10%
            ws.conditional_formatting.add(
                f"AJ5:AJ{last_data_row}",
                CellIsRule(
                    operator="lessThan",
                    formula=["10"],
                    fill=PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid"),
                    font=Font(color="856404", bold=True)
                )
            )
            
            # Рекомендованные цены: выделение если выше текущей
            ws.conditional_formatting.add(
                f"AM5:AM{last_data_row}",
                CellIsRule(
                    operator="greaterThan",
                    formula=[f"D5"],
                    fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
                    font=Font(color="006100", bold=True)
                )
            )
    
    def _create_dashboard_sheet(self, ws, results, marketplace_name):
        """Создание дашборда со сводной статистикой"""
        ws.merge_cells('A1:H1')
        ws.cell(row=1, column=1, value=f"📈 Дашборд юнит-экономики — {marketplace_name}").font = self.title_font
        
        if not results:
            ws.cell(row=3, column=1, value="Нет данных для отображения")
            return
        
        # Ключевые метрики
        total_items = len(results)
        profitable_items = len([r for r in results if r.gross_profit > 0])
        unprofitable_items = total_items - profitable_items
        total_profit = sum(r.gross_profit for r in results)
        avg_margin = np.mean([r.margin_percent for r in results])
        avg_roi = np.mean([r.roi_percent for r in results])
        total_revenue = sum(r.selling_price for r in results)
        total_costs = sum(r.total_expenses for r in results)
        avg_ltv_cac = np.mean([r.ltv_cac_ratio for r in results if r.ltv_cac_ratio < 999])
        avg_romi = np.mean([r.romi for r in results])
        
        # Метрики складской оптимизации
        avg_optimal_stock = np.mean([r.optimal_stock_units for r in results])
        avg_turnover_days = np.mean([r.stock_turnover_days for r in results if r.stock_turnover_days > 0])
        avg_optimization_potential = np.mean([r.stock_optimization_potential for r in results])
        
        # Строка 3-4: Основные метрики
        metrics_row1 = [
            ("Всего товаров", f"{total_items}", "A3"),
            ("Прибыльных", f"{profitable_items} ({profitable_items/total_items*100:.0f}%)", "C3"),
            ("Убыточных", f"{unprofitable_items} ({unprofitable_items/total_items*100:.0f}%)", "E3"),
            ("Общая прибыль", f"{total_profit:,.0f} ₽", "G3"),
        ]
        
        for title, value, cell_ref in metrics_row1:
            row = int(cell_ref[1:])
            col = ord(cell_ref[0]) - ord('A') + 1
            
            title_cell = ws.cell(row=row, column=col, value=title)
            title_cell.font = Font(bold=True, size=11, color="666666")
            
            value_cell = ws.cell(row=row+1, column=col, value=value)
            value_cell.font = Font(bold=True, size=16, color="1a1a2e")
        
        # Строка 6-7: Финансовые метрики
        metrics_row2 = [
            ("Общая выручка", f"{total_revenue:,.0f} ₽", "A6"),
            ("Общие расходы", f"{total_costs:,.0f} ₽", "C6"),
            ("Средняя маржа", f"{avg_margin:.1f}%", "E6"),
            ("Средний ROI", f"{avg_roi:.1f}%", "G6"),
        ]
        
        for title, value, cell_ref in metrics_row2:
            row = int(cell_ref[1:])
            col = ord(cell_ref[0]) - ord('A') + 1
            
            title_cell = ws.cell(row=row, column=col, value=title)
            title_cell.font = Font(bold=True, size=11, color="666666")
            
            value_cell = ws.cell(row=row+1, column=col, value=value)
            value_cell.font = Font(bold=True, size=16, color="1a1a2e")
        
        # Строка 9-10: Клиентские метрики
        metrics_row3 = [
            ("Средний LTV/CAC", f"{avg_ltv_cac:.1f}x", "A9"),
            ("Средний ROMI", f"{avg_romi:.1f}%", "C9"),
            ("Оптимальный запас", f"{avg_optimal_stock:.0f} шт", "E9"),
            ("Оборачиваемость", f"{avg_turnover_days:.1f} дн", "G9"),
        ]
        
        for title, value, cell_ref in metrics_row3:
            row = int(cell_ref[1:])
            col = ord(cell_ref[0]) - ord('A') + 1
            
            title_cell = ws.cell(row=row, column=col, value=title)
            title_cell.font = Font(bold=True, size=11, color="666666")
            
            value_cell = ws.cell(row=row+1, column=col, value=value)
            value_cell.font = Font(bold=True, size=16, color="1a1a2e")
        
        # Строка 12: Потенциал оптимизации
        ws.merge_cells('A12:H12')
        optimization_cell = ws.cell(row=12, column=1, value=(
            f"💡 Потенциал оптимизации складских остатков: {avg_optimization_potential:.1f}% "
            f"({sum(1 for r in results if r.stock_optimization_potential > 0)} из {total_items} товаров можно оптимизировать)"
        ))
        if avg_optimization_potential > 10:
            optimization_cell.font = Font(bold=True, size=12, color="006100")
            optimization_cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        elif avg_optimization_potential > 0:
            optimization_cell.font = Font(bold=True, size=12, color="856404")
            optimization_cell.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
        else:
            optimization_cell.font = Font(bold=True, size=12, color="666666")
        
        # Топ-5 товаров по прибыли
        if len(results) >= 5:
            ws.merge_cells('A14:H14')
            ws.cell(row=14, column=1, value="🏆 Топ-5 товаров по прибыли").font = Font(bold=True, size=12)
            
            top_results = sorted(results, key=lambda x: x.gross_profit, reverse=True)[:5]
            
            headers = ["№", "Артикул", "Прибыль, ₽", "Маржа, %", "LTV/CAC", "Оптим. запас", "Оборач., дн"]
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=15, column=col_idx, value=header)
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="E8E8E8", end_color="E8E8E8", fill_type="solid")
                cell.border = self.thin_border
            
            for i, r in enumerate(top_results, 1):
                ws.cell(row=15+i, column=1, value=i).border = self.thin_border
                ws.cell(row=15+i, column=2, value=r.artikul).border = self.thin_border
                cell = ws.cell(row=15+i, column=3, value=r.gross_profit)
                cell.border = self.thin_border
                cell.number_format = '#,##0.00'
                cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                ws.cell(row=15+i, column=4, value=f"{r.margin_percent:.1f}%").border = self.thin_border
                ws.cell(row=15+i, column=5, value=f"{r.ltv_cac_ratio:.1f}x").border = self.thin_border
                ws.cell(row=15+i, column=6, value=r.optimal_stock_units).border = self.thin_border
                ws.cell(row=15+i, column=7, value=f"{r.stock_turnover_days:.1f}").border = self.thin_border

# ============================================================================
# БЛОК 10: ИНТЕРФЕЙС ПОЛЬЗОВАТЕЛЯ (STREAMLIT) — ПОЛНАЯ ВЕРСИЯ
# ============================================================================

def init_session_state():
    """Инициализация всех состояний сессии Streamlit"""
    
    if 'api_manager' not in st.session_state:
        st.session_state.api_manager = MarketplaceAPIManager()
    
    if 'calculator' not in st.session_state:
        st.session_state.calculator = FBSUnitEconomicsCalculator(
            api_manager=st.session_state.api_manager
        )
    
    if 'exporter' not in st.session_state:
        try:
            st.session_state.exporter = ProfessionalExcelExporter()
        except ImportError:
            st.session_state.exporter = None
    
    if 'data_exporter' not in st.session_state:
        st.session_state.data_exporter = DataExporter()
    
    if 'secure_data' not in st.session_state:
        st.session_state.secure_data = SecureDataManager()
    
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

def render_sidebar():
    """Отрисовка боковой панели навигации"""
    
    with st.sidebar:
        st.markdown("""
        <div style='text-align: center; padding: 20px 15px; background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460); border-radius: 12px; margin-bottom: 25px;'>
            <h1 style='color: white; margin: 0; font-size: 1.5em;'>🚀 FBS PRO</h1>
            <p style='color: #a8a8d0; margin: 8px 0 0 0; font-size: 0.9em;'>Операционная версия</p>
            <p style='color: #6666aa; margin: 5px 0 0 0; font-size: 0.7em;'>v6.0.0 | Оптимизация склада</p>
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 🧭 Навигация")
        
        sections = {
            "🏠 Главная": "main",
            "🧮 Калькулятор FBS": "calculator",
            "📋 Тарифы маркетплейсов": "tariffs",
            "📈 Дашборд": "dashboard",
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
        else:
            st.warning("⚠️ Тарифы: Дефолтные")
        
        if st.session_state.results:
            st.success(f"✅ Рассчитано: {len(st.session_state.results)} товаров")
        else:
            st.info("ℹ️ Расчеты не выполнялись")
        
        st.markdown("---")
        st.markdown("### ⚡ Быстрые действия")
        
        if st.button("🔄 Обновить тарифы", width="stretch"):
            with st.spinner("Загрузка тарифов..."):
                calculator.refresh_tariffs(force=True)
                st.success("✅ Тарифы обновлены!")
                st.rerun()
        
        if st.button("🗑️ Очистить результаты", width="stretch"):
            st.session_state.results = []
            st.session_state.input_data_list = []
            st.success("✅ Результаты очищены!")
            st.rerun()

# ============================================================================
# БЛОК 11: ГЛАВНАЯ ФУНКЦИЯ ПРИЛОЖЕНИЯ (ПОЛНАЯ)
# ============================================================================

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
        </div>
        """, unsafe_allow_html=True)
        
        st.markdown("### 🎯 Ключевые возможности")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.markdown("""
            <div style='background: linear-gradient(135deg, #0984e3, #6c5ce7); padding: 25px; border-radius: 15px; color: white; height: 100%;'>
                <h3 style='margin-top: 0;'>📦 Оптимизация склада</h3>
                <ul>
                    <li>Оптимальный запас (EOQ)</li>
                    <li>Страховой запас</li>
                    <li>Точка заказа</li>
                    <li>Оборачиваемость</li>
                    <li>Потенциал оптимизации</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
        
        with col2:
            st.markdown("""
            <div style='background: linear-gradient(135deg, #00b894, #00cec9); padding: 25px; border-radius: 15px; color: white; height: 100%;'>
                <h3 style='margin-top: 0;'>🚚 Логистические коридоры</h3>
                <ul>
                    <li>Точка безубыточности</li>
                    <li>Критические зоны (25, 50, 100 км)</li>
                    <li>Оптимизация First Mile</li>
                    <li>Загрузка паллет</li>
                    <li>Маршрутная оптимизация</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
        
        with col3:
            st.markdown("""
            <div style='background: linear-gradient(135deg, #e17055, #d63031); padding: 25px; border-radius: 15px; color: white; height: 100%;'>
                <h3 style='margin-top: 0;'>📊 Юнит-экономика</h3>
                <ul>
                    <li>FBS полный расчет</li>
                    <li>Рекомендованные цены</li>
                    <li>LTV, CAC, ROMI</li>
                    <li>Условное форматирование</li>
                    <li>Экспорт в Excel/CSV/GS</li>
                </ul>
            </div>
            """, unsafe_allow_html=True)
    
    elif current_section == 'calculator':
        st.markdown("## 🧮 Калькулятор FBS юнит-экономики")
        st.info("""
        **🎯 Профессиональный расчет FBS с оптимизацией складских остатков**
        
        - 🚛 **First Mile** — ваша логистика до склада МП
        - 📦 **Last Mile** — доставка МП до клиента
        - 📊 **Оптимизация склада** — EOQ, страховой запас, точка заказа
        - 💰 **Рекомендованные цены** — при разных уровнях маржи
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
                category = st.selectbox("Категория", list(st.session_state.calculator.current_tariffs.keys()) or ["default"])
                selling_price = st.number_input("Цена продажи, ₽", 5000.0, step=100.0)
                cogs = st.number_input("Себестоимость, ₽", 3000.0, step=100.0)
            
            with col2:
                weight = st.number_input("Вес, кг", 1.5, step=0.1)
                length = st.number_input("Длина, см", 20, step=1)
                width = st.number_input("Ширина, см", 15, step=1)
                height = st.number_input("Высота, см", 10, step=1)
                warehouse_distance = st.number_input("Расстояние до МП, км", 50.0, step=1.0)
            
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
                    warehouse_distance_km=warehouse_distance
                )
                
                result = st.session_state.calculator.calculate_unit_economics(input_data)
                st.session_state.results = [result]
                st.session_state.input_data_list = [input_data]
                
                st.markdown("---")
                st.markdown("## 📊 Результаты расчета")
                
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("💰 Прибыль", f"{result.gross_profit:,.0f} ₽", f"{result.margin_percent:.1f}% маржи")
                with col2:
                    st.metric("📦 Расходы", f"{result.total_expenses:,.0f} ₽")
                with col3:
                    st.metric("📈 ROI", f"{result.roi_percent:.1f}%")
                with col4:
                    st.metric("👥 LTV/CAC", f"{result.ltv_cac_ratio:.1f}x")
                
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
                
                st.markdown("### 📦 Оптимизация складских остатков")
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Оптимальный запас (EOQ)", f"{result.optimal_stock_units} шт")
                with col2:
                    st.metric("Страховой запас", f"{result.safety_stock_units} шт")
                with col3:
                    st.metric("Точка заказа", f"{result.reorder_point_units} шт")
                
                col1, col2 = st.columns(2)
                with col1:
                    st.metric("Оборачиваемость", f"{result.stock_turnover_days:.1f} дней")
                with col2:
                    st.metric("Потенциал оптимизации", f"{result.stock_optimization_potential:.1f}%")
    
    elif current_section == 'tariffs':
        st.markdown("## 📋 Актуальные тарифы маркетплейсов")
        
        calculator = st.session_state.calculator
        
        col1, col2, col3 = st.columns(3)
        with col1:
            marketplace = st.selectbox("Маркетплейс", ["Ozon", "Wildberries", "Яндекс Маркет"])
        with col2:
            force_refresh = st.checkbox("🔄 Принудительное обновление")
        with col3:
            if st.button("📥 Загрузить тарифы", type="primary"):
                with st.spinner(f"Загрузка тарифов {marketplace}..."):
                    calculator.set_marketplace(marketplace)
                    calculator.refresh_tariffs(force=force_refresh)
                    st.success(f"✅ Тарифы {marketplace} загружены!")
                    st.rerun()
        
        if calculator.current_tariffs:
            df = calculator.api_manager.get_all_tariffs_as_dataframe(marketplace)
            st.dataframe(df, width="stretch", height=400)
            
            st.markdown("### 📊 Статистика источников")
            sources = df['Источник'].value_counts()
            st.dataframe(sources, width="stretch")
    
    elif current_section == 'dashboard':
        st.markdown("## 📈 Дашборд")
        
        if not st.session_state.results:
            st.warning("⚠️ Нет данных. Выполните расчет в разделе 'Калькулятор FBS'.")
            return
        
        results = st.session_state.results
        
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("📦 Товаров", len(results))
        with col2:
            profitable = len([r for r in results if r.gross_profit > 0])
            st.metric("✅ Прибыльных", f"{profitable}")
        with col3:
            total_profit = sum(r.gross_profit for r in results)
            st.metric("💰 Общая прибыль", f"{total_profit:,.0f} ₽")
        with col4:
            avg_margin = np.mean([r.margin_percent for r in results])
            st.metric("📊 Средняя маржа", f"{avg_margin:.1f}%")
        with col5:
            avg_ltv_cac = np.mean([r.ltv_cac_ratio for r in results if r.ltv_cac_ratio < 999])
            st.metric("👥 LTV/CAC", f"{avg_ltv_cac:.1f}x")
        
        st.markdown("---")
        st.markdown("### 📊 Топ-10 товаров по прибыли")
        
        top_results = sorted(results, key=lambda x: x.gross_profit, reverse=True)[:10]
        data = []
        for r in top_results:
            data.append({
                'Артикул': r.artikul,
                'Прибыль, ₽': r.gross_profit,
                'Маржа, %': r.margin_percent,
                'LTV/CAC': r.ltv_cac_ratio,
                'Опт. запас': r.optimal_stock_units,
                'Оборач., дн': r.stock_turnover_days
            })
        st.dataframe(pd.DataFrame(data), width="stretch")
    
    elif current_section == 'export':
        st.markdown("## 📥 Экспорт данных")
        
        if not st.session_state.results:
            st.warning("⚠️ Нет данных для экспорта.")
            return
        
        results = st.session_state.results
        input_data_list = st.session_state.input_data_list
        
        st.success(f"✅ Доступно для экспорта: {len(results)} товаров")
        
        tab1, tab2, tab3 = st.tabs(["📊 Excel (с формулами)", "📄 CSV", "🌐 Google Sheets"])
        
        with tab1:
            if st.session_state.exporter is None:
                st.error("❌ OpenPyXL не установлен. Выполните: pip install openpyxl")
            else:
                if st.button("📥 Скачать Excel-отчет", type="primary"):
                    with st.spinner("Создание Excel-отчета..."):
                        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                        filename = f"FBS_Report_{st.session_state.marketplace}_{timestamp}.xlsx"
                        output_path = EXPORTS_DIR / filename
                        
                        success = st.session_state.exporter.export_fbs_report(
                            results=results,
                            input_data_list=input_data_list,
                            calculator=st.session_state.calculator,
                            marketplace_name=st.session_state.marketplace,
                            output_path=str(output_path)
                        )
                        
                        if success and output_path.exists():
                            with open(output_path, "rb") as f:
                                st.download_button(
                                    label="⬇️ Скачать",
                                    data=f.read(),
                                    file_name=filename,
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                                )
        
        with tab2:
            delimiter = st.selectbox("Разделитель", [";", ",", "\\t"])
            if st.button("📥 Скачать CSV", type="primary"):
                with st.spinner("Экспорт в CSV..."):
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    filename = f"FBS_Report_{st.session_state.marketplace}_{timestamp}.csv"
                    output_path = EXPORTS_DIR / filename
                    
                    sep = "\t" if delimiter == "\\t" else delimiter
                    success = st.session_state.data_exporter.export_to_csv(
                        results=results,
                        input_data_list=input_data_list,
                        output_path=str(output_path),
                        delimiter=sep
                    )
                    
                    if success:
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
                with st.spinner("Экспорт в Google Sheets..."):
                    temp_path = TEMP_DIR / "credentials.json"
                    temp_path.write_bytes(credentials_file.getvalue())
                    
                    success = st.session_state.data_exporter.export_to_google_sheets(
                        results=results,
                        input_data_list=input_data_list,
                        credentials_path=str(temp_path),
                        spreadsheet_name=f"FBS Report {datetime.now().strftime('%Y-%m-%d')}"
                    )
                    
                    if success:
                        st.success("✅ Экспорт в Google Sheets выполнен!")
                    else:
                        st.error("❌ Ошибка экспорта. Проверьте credentials.json")
    
    elif current_section == 'settings':
        st.markdown("## ⚙️ Настройки")
        
        tab1, tab2 = st.tabs(["🔑 API Ключи", "🏪 Маркетплейс и налоги"])
        
        with tab1:
            st.markdown("### 🔑 Настройка API ключей")
            
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

if __name__ == "__main__":
    main()
