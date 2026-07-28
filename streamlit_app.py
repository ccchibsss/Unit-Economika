"""
============================================================================
🚀 FBS UNIT ECONOMICS PRO 2026 — ПОЛНАЯ ВЕРСИЯ С API ИНТЕГРАЦИЕЙ
============================================================================
Профессиональный калькулятор юнит-экономики для FBS-модели
Маркетплейсы: Ozon, Wildberries, Яндекс Маркет
Версия: 5.0.0

ОСНОВНЫЕ ВОЗМОЖНОСТИ:
- Автоматическая загрузка тарифов через API маркетплейсов
- AI-обогащение данных через DeepSeek API
- Отдельный лист "Тарифы МП" в Excel с живыми ссылками
- Формулы тянут данные с листа тарифов
- Кэширование API-запросов
- Автообновление тарифов
- Полный расчет юнит-экономики FBS (First Mile + Last Mile)
- Расчет штрафов за просрочку (Penalty Rate)
- Стоимость обработки заказа (Pick & Pack)
- Точка безубыточности по расстоянию (First Mile)
- Расчет LTV и CAC с адаптацией под FBS
- Запас прочности по цене для сезонных распродаж
- Анализ перехода на FBO/FBP
- Экспорт в Excel с живыми формулами
- Экспорт в Google Sheets
- Интерактивная визуализация
- Поддержка до 500 000+ товаров через батчевую обработку
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

APP_VERSION = "5.0.0"
APP_NAME = "🚀 FBS Юнит-экономика PRO 2026"
APP_DESCRIPTION = "Профессиональный расчет юнит-экономики для FBS-модели с автообновлением тарифов через API"

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
# БЛОК 6: ДАТАКЛАССЫ ДЛЯ РАСЧЕТОВ
# ============================================================================

@dataclass
class FBSInputData:
    """
    Входные данные для расчета FBS юнит-экономики.
    Содержит все необходимые параметры товара и бизнеса.
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
    
    # Складские параметры
    stock_depth_days: int = 30  # Глубина складского запаса в днях
    daily_sales: int = 5  # Среднее количество продаж в день
    warehouse_rent_per_sqm: float = 500.0  # Стоимость аренды склада за м² в месяц
    warehouse_space_per_unit: float = 0.01  # Занимаемая площадь на складе на единицу (м²)
    
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
        # Фильтруем только известные поля
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
        
        return errors

@dataclass
class FBSResultData:
    """
    Результаты расчета FBS юнит-экономики.
    Содержит полную детализацию всех расходов и метрик.
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
    
    # Сравнение с другими моделями
    fbo_profit: float = 0.0  # Прибыль при модели FBO
    fbp_profit: float = 0.0  # Прибыль при модели FBP
    recommended_model: str = ""  # Рекомендуемая модель (FBS/FBO/FBP)
    
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
            'recommended_model': self.recommended_model,
            'first_mile_cost': self.first_mile_cost,
            'last_mile_cost': self.last_mile_cost,
            'penalty_cost': self.penalty_cost,
            'ltv': self.ltv,
            'cac': self.cac,
            'ltv_cac_ratio': self.ltv_cac_ratio
        }

# ============================================================================
# БЛОК 7: ОСНОВНОЙ КАЛЬКУЛЯТОР FBS ЮНИТ-ЭКОНОМИКИ
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
    - Использует динамически загружаемые тарифы через API
    - Учитывает специфику FBS: двойную логистику, штрафы, Pick & Pack
    - Рассчитывает LTV, CAC, точку безубыточности по расстоянию
    - Сравнивает модели FBS, FBO, FBP
    - Поддерживает пакетную обработку до 500 000+ товаров
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
        12. LTV и CAC
        13. Сравнение с FBO и FBP
        
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
        
        # Получаем актуальный тариф для категории
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
        
        # =====================================================================
        # 16. СРАВНЕНИЕ С МОДЕЛЯМИ FBO И FBP
        # =====================================================================
        fbo_multiplier = tariff.get('fbo_multiplier', 0.75)
        fbp_multiplier = tariff.get('fbp_multiplier', 0.60)
        storage_base_rate = tariff.get('storage_base_rate', 0.30)
        
        # Расчет стоимости хранения при FBO
        if input_data.length_cm > 0 and input_data.width_cm > 0 and input_data.height_cm > 0:
            storage_volume_liters = (input_data.length_cm * input_data.width_cm * input_data.height_cm) / 1000.0
        else:
            storage_volume_liters = 5.0  # Средний объем если габариты не указаны
        
        fbo_storage_days = 30  # Среднее время хранения на складе МП
        fbo_storage_cost = storage_volume_liters * storage_base_rate * fbo_storage_days
        
        # FBO: экономим на First Mile, но платим за хранение
        fbo_commission = result.commission  # Комиссия такая же
        fbo_logistics = result.last_mile_cost * fbo_multiplier  # Логистика дешевле
        
        fbo_total_expenses = (
            input_data.cogs +
            fbo_commission +
            fbo_logistics +
            fbo_storage_cost +
            result.acquiring_cost +
            result.return_cost +
            result.tax_cost +
            result.marketing_cost +
            result.packaging_cost
        )
        
        result.fbo_profit = input_data.selling_price - fbo_total_expenses
        
        # FBP: частичный фулфилмент
        fbp_logistics = result.last_mile_cost * fbp_multiplier
        fbp_storage_cost = fbo_storage_cost * 0.5  # Меньше времени хранения
        
        fbp_total_expenses = (
            input_data.cogs +
            fbo_commission +
            fbp_logistics +
            fbp_storage_cost +
            result.acquiring_cost +
            result.return_cost +
            result.tax_cost +
            result.marketing_cost +
            result.packaging_cost
        )
        
        result.fbp_profit = input_data.selling_price - fbp_total_expenses
        
        # Рекомендация оптимальной модели
        profits = {
            "FBS": result.gross_profit,
            "FBO": result.fbo_profit,
            "FBP": result.fbp_profit
        }
        
        result.recommended_model = max(profits, key=profits.get)
        
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
# БЛОК 8: ВИЗУАЛИЗАЦИЯ И ГРАФИКИ
# ============================================================================

class FBSVisualizer:
    """
    Класс для создания профессиональных интерактивных визуализаций.
    Использует Plotly для построения графиков.
    """
    
    # Цветовая палитра приложения
    COLORS = {
        'primary': '#1a1a2e',
        'secondary': '#16213e',
        'accent': '#0f3460',
        'highlight': '#e94560',
        'success': '#00b894',
        'warning': '#fdcb6e',
        'danger': '#d63031',
        'info': '#0984e3',
        'light': '#dfe6e9',
        'dark': '#2d3436',
        'purple': '#6c5ce7',
        'teal': '#00cec9',
        'orange': '#e17055',
        'gradient_1': ['#00b894', '#00cec9', '#0984e3', '#6c5ce7', '#a29bfe'],
        'gradient_2': ['#d63031', '#e17055', '#fdcb6e', '#00b894', '#0984e3'],
        'gradient_3': ['#1a1a2e', '#16213e', '#0f3460', '#0984e3', '#00b894'],
        'fbs_color': '#0984e3',
        'fbo_color': '#00b894',
        'fbp_color': '#6c5ce7'
    }
    
    @staticmethod
    def create_cost_breakdown_pie(result: FBSResultData, 
                                  title: str = "Структура расходов FBS") -> go.Figure:
        """
        Создание круговой диаграммы структуры расходов.
        
        Args:
            result: Результаты расчета
            title: Заголовок диаграммы
            
        Returns:
            go.Figure: Интерактивная круговая диаграмма
        """
        # Формируем категории расходов
        cost_categories = {}
        
        # Себестоимость (вычисляем как разницу)
        cogs_value = result.total_expenses - (
            result.commission + result.first_mile_cost + result.last_mile_cost +
            result.pick_pack_cost + result.packaging_cost + result.acquiring_cost +
            result.return_cost + result.penalty_cost + result.marketing_cost +
            result.warehouse_cost + result.tax_cost
        )
        
        if cogs_value > 0.01:
            cost_categories['Себестоимость закупки'] = cogs_value
        
        # Добавляем все статьи расходов
        expense_items = [
            ('Комиссия маркетплейса', result.commission),
            ('First Mile (доставка до МП)', result.first_mile_cost),
            ('Last Mile (доставка клиенту)', result.last_mile_cost),
            ('Pick & Pack (обработка заказа)', result.pick_pack_cost),
            ('Упаковочные материалы', result.packaging_cost),
            ('Эквайринг', result.acquiring_cost),
            ('Возвраты', result.return_cost),
            ('Штрафы за просрочку', result.penalty_cost),
            ('Маркетинговые расходы', result.marketing_cost),
            ('Складские расходы', result.warehouse_cost),
            ('Налоги', result.tax_cost)
        ]
        
        for name, value in expense_items:
            if value > 0.01:
                cost_categories[name] = value
        
        # Сортируем по убыванию для лучшей читаемости
        cost_categories = dict(sorted(cost_categories.items(), key=lambda x: x[1], reverse=True))
        
        # Создаем диаграмму
        fig = go.Figure(data=[go.Pie(
            labels=list(cost_categories.keys()),
            values=list(cost_categories.values()),
            hole=0.45,  # Делаем donut chart
            marker=dict(
                colors=FBSVisualizer.COLORS['gradient_1'][:len(cost_categories)],
                line=dict(color='white', width=2)
            ),
            textinfo='label+percent',
            textposition='outside',
            textfont=dict(size=11, family='Arial'),
            hovertemplate=(
                '<b>%{label}</b><br>' +
                'Сумма: %{value:,.2f} ₽<br>' +
                'Доля: %{percent}<br>' +
                '<extra></extra>'
            ),
            pull=[0.05] * len(cost_categories)  # Небольшое разделение секторов
        )])
        
        # Настройка макета
        fig.update_layout(
            title=dict(
                text=f"<b>{title}</b><br>" +
                     f"<sub>Общие расходы: {result.total_expenses:,.2f} ₽ | " +
                     f"Прибыль: {result.gross_profit:,.2f} ₽</sub>",
                font=dict(size=18, color=FBSVisualizer.COLORS['primary'], family='Arial'),
                x=0.5,
                xanchor='center'
            ),
            template="plotly_white",
            height=550,
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.3,
                xanchor="center",
                x=0.5,
                font=dict(size=10)
            ),
            margin=dict(t=120, b=120, l=20, r=20),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        
        return fig
    
    @staticmethod
    def create_waterfall_chart(result: FBSResultData) -> go.Figure:
        """
        Создание водопадной диаграммы формирования прибыли.
        Показывает как цена превращается в чистую прибыль после всех вычетов.
        
        Args:
            result: Результаты расчета
            
        Returns:
            go.Figure: Водопадная диаграмма
        """
        # Категории и значения для водопадной диаграммы
        categories = [
            "Цена продажи",
            "Себестоимость",
            "Комиссия МП",
            "First Mile",
            "Last Mile",
            "Pick & Pack",
            "Упаковка",
            "Эквайринг",
            "Возвраты",
            "Штрафы",
            "Маркетинг",
            "Склад",
            "Налог",
            "ЧИСТАЯ ПРИБЫЛЬ"
        ]
        
        # Вычисляем себестоимость
        cogs_value = result.total_expenses - (
            result.commission + result.first_mile_cost + result.last_mile_cost +
            result.pick_pack_cost + result.packaging_cost + result.acquiring_cost +
            result.return_cost + result.penalty_cost + result.marketing_cost +
            result.warehouse_cost + result.tax_cost
        )
        
        values = [
            result.selling_price,      # Начальная точка
            -cogs_value,                # Вычитаем себестоимость
            -result.commission,         # Вычитаем комиссию
            -result.first_mile_cost,    # Вычитаем First Mile
            -result.last_mile_cost,     # Вычитаем Last Mile
            -result.pick_pack_cost,     # Вычитаем Pick & Pack
            -result.packaging_cost,     # Вычитаем упаковку
            -result.acquiring_cost,     # Вычитаем эквайринг
            -result.return_cost,        # Вычитаем возвраты
            -result.penalty_cost,       # Вычитаем штрафы
            -result.marketing_cost,     # Вычитаем маркетинг
            -result.warehouse_cost,     # Вычитаем склад
            -result.tax_cost,           # Вычитаем налог
            result.gross_profit         # Итоговая прибыль
        ]
        
        # Определяем цвета
        total_color = FBSVisualizer.COLORS['success'] if result.gross_profit > 0 else FBSVisualizer.COLORS['danger']
        
        # Создаем водопадную диаграмму
        fig = go.Figure(data=[go.Waterfall(
            name="Формирование прибыли",
            orientation="v",
            measure=["absolute"] + ["relative"] * 12 + ["total"],
            x=categories,
            y=values,
            text=[f"{v:,.0f} ₽" for v in values],
            textposition="outside",
            textfont=dict(size=10),
            connector={"line": {"color": "rgb(63, 63, 63)", "width": 1}},
            increasing={"marker": {"color": FBSVisualizer.COLORS['success']}},
            decreasing={"marker": {"color": FBSVisualizer.COLORS['danger']}},
            totals={"marker": {"color": total_color, "line": {"color": total_color, "width": 2}}}
        )])
        
        # Добавляем горизонтальную линию нуля
        fig.add_hline(
            y=0,
            line_dash="dash",
            line_color="gray",
            line_width=1,
            opacity=0.5
        )
        
        # Настройка макета
        fig.update_layout(
            title=dict(
                text="<b>Водопадная диаграмма формирования прибыли</b><br>" +
                     f"<sub>От цены {result.selling_price:,.0f} ₽ до чистой прибыли {result.gross_profit:,.0f} ₽</sub>",
                font=dict(size=18, color=FBSVisualizer.COLORS['primary'], family='Arial')
            ),
            template="plotly_white",
            height=550,
            showlegend=False,
            xaxis=dict(
                tickangle=45,
                tickfont=dict(size=10)
            ),
            yaxis=dict(
                title="Сумма, ₽",
                gridcolor='rgba(0,0,0,0.1)'
            ),
            margin=dict(t=120, b=100, l=80, r=20),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        
        return fig
    
    @staticmethod
    def create_models_comparison_chart(results: List[FBSResultData]) -> go.Figure:
        """
        Создание сравнительной диаграммы моделей FBS vs FBO vs FBP.
        
        Args:
            results: Список результатов расчета
            
        Returns:
            go.Figure: Групповая столбчатая диаграмма
        """
        if not results:
            fig = go.Figure()
            fig.add_annotation(
                text="Нет данных для отображения",
                showarrow=False,
                font=dict(size=16)
            )
            return fig
        
        # Берем топ-15 товаров по прибыли FBS
        top_results = sorted(results, key=lambda x: x.gross_profit, reverse=True)[:15]
        
        artikuls = [r.artikul[:20] if len(r.artikul) > 20 else r.artikul for r in top_results]
        fbs_profits = [r.gross_profit for r in top_results]
        fbo_profits = [r.fbo_profit for r in top_results]
        fbp_profits = [r.fbp_profit for r in top_results]
        
        fig = go.Figure()
        
        # FBS столбцы
        fig.add_trace(go.Bar(
            name='FBS (со своего склада)',
            x=artikuls,
            y=fbs_profits,
            marker_color=FBSVisualizer.COLORS['fbs_color'],
            text=[f'{v:,.0f} ₽' for v in fbs_profits],
            textposition='auto',
            textfont=dict(size=9),
            hovertemplate='<b>FBS</b><br>%{x}<br>Прибыль: %{y:,.0f} ₽<extra></extra>'
        ))
        
        # FBO столбцы
        fig.add_trace(go.Bar(
            name='FBO (со склада МП)',
            x=artikuls,
            y=fbo_profits,
            marker_color=FBSVisualizer.COLORS['fbo_color'],
            text=[f'{v:,.0f} ₽' for v in fbo_profits],
            textposition='auto',
            textfont=dict(size=9),
            hovertemplate='<b>FBO</b><br>%{x}<br>Прибыль: %{y:,.0f} ₽<extra></extra>'
        ))
        
        # FBP столбцы
        fig.add_trace(go.Bar(
            name='FBP (частичный фулфилмент)',
            x=artikuls,
            y=fbp_profits,
            marker_color=FBSVisualizer.COLORS['fbp_color'],
            text=[f'{v:,.0f} ₽' for v in fbp_profits],
            textposition='auto',
            textfont=dict(size=9),
            hovertemplate='<b>FBP</b><br>%{x}<br>Прибыль: %{y:,.0f} ₽<extra></extra>'
        ))
        
        # Линия нуля
        fig.add_hline(
            y=0,
            line_dash="solid",
            line_color="red",
            line_width=1,
            opacity=0.3
        )
        
        # Настройка макета
        fig.update_layout(
            title=dict(
                text="<b>Сравнение моделей фулфилмента: FBS vs FBO vs FBP</b><br>" +
                     "<sub>Топ-15 товаров по прибыли FBS</sub>",
                font=dict(size=18, color=FBSVisualizer.COLORS['primary'])
            ),
            barmode='group',
            template="plotly_white",
            height=500,
            xaxis=dict(
                tickangle=45,
                tickfont=dict(size=9),
                title="Артикул товара"
            ),
            yaxis=dict(
                title="Прибыль, ₽",
                gridcolor='rgba(0,0,0,0.1)'
            ),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1,
                font=dict(size=11)
            ),
            margin=dict(t=120, b=100, l=80, r=20),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        
        return fig
    
    @staticmethod
    def create_break_even_distance_chart(results: List[FBSResultData]) -> go.Figure:
        """
        Визуализация точки безубыточности по расстоянию First Mile.
        
        Args:
            results: Список результатов расчета
            
        Returns:
            go.Figure: Горизонтальная столбчатая диаграмма
        """
        if not results:
            fig = go.Figure()
            fig.add_annotation(text="Нет данных", showarrow=False)
            return fig
        
        # Фильтруем товары с конечной точкой безубыточности
        valid_results = [
            r for r in results 
            if r.break_even_distance_km < 999999 and r.break_even_distance_km > 0
        ]
        
        # Сортируем по возрастанию (самые проблемные первые)
        valid_results = sorted(valid_results, key=lambda x: x.break_even_distance_km)[:20]
        
        if not valid_results:
            fig = go.Figure()
            fig.add_annotation(
                text="Все товары имеют большой запас по расстоянию",
                showarrow=False,
                font=dict(size=14)
            )
            return fig
        
        artikuls = [r.artikul[:25] for r in valid_results]
        distances = [r.break_even_distance_km for r in valid_results]
        profits = [r.gross_profit for r in valid_results]
        
        # Цветовая шкала: красный - опасно, зеленый - хорошо
        colors = []
        for d in distances:
            if d < 25:
                colors.append(FBSVisualizer.COLORS['danger'])
            elif d < 50:
                colors.append(FBSVisualizer.COLORS['warning'])
            elif d < 100:
                colors.append(FBSVisualizer.COLORS['info'])
            else:
                colors.append(FBSVisualizer.COLORS['success'])
        
        fig = go.Figure(data=[go.Bar(
            y=artikuls,
            x=distances,
            orientation='h',
            marker=dict(
                color=colors,
                line=dict(color='white', width=1)
            ),
            text=[f'{d:.1f} км (прибыль {p:,.0f} ₽)' for d, p in zip(distances, profits)],
            textposition='auto',
            textfont=dict(size=9),
            hovertemplate=(
                '<b>%{y}</b><br>' +
                'Макс. расстояние: %{x:.1f} км<br>' +
                'Прибыль: %{customdata:,.0f} ₽<br>' +
                '<extra></extra>'
            ),
            customdata=profits
        )])
        
        # Критические зоны
        fig.add_vline(
            x=25, line_dash="dash", line_color="red", line_width=2,
            annotation_text="Критическая зона (< 25 км)",
            annotation_position="top right",
            annotation_font=dict(size=10, color="red")
        )
        
        fig.add_vline(
            x=50, line_dash="dash", line_color="orange", line_width=2,
            annotation_text="Зона риска (< 50 км)",
            annotation_position="top right",
            annotation_font=dict(size=10, color="orange")
        )
        
        # Настройка макета
        fig.update_layout(
            title=dict(
                text="<b>Точка безубыточности по расстоянию First Mile</b><br>" +
                     "<sub>Максимальное расстояние до склада МП для безубыточной работы</sub>",
                font=dict(size=18, color=FBSVisualizer.COLORS['primary'])
            ),
            template="plotly_white",
            height=550,
            xaxis=dict(
                title="Максимальное расстояние, км",
                gridcolor='rgba(0,0,0,0.1)'
            ),
            yaxis=dict(
                title="Товар",
                tickfont=dict(size=9)
            ),
            margin=dict(t=120, b=50, l=200, r=50),
            paper_bgcolor='rgba(0,0,0,0)',
            plot_bgcolor='rgba(0,0,0,0)'
        )
        
        return fig
    
    @staticmethod
    def create_ltv_cac_gauge(result: FBSResultData) -> go.Figure:
        """
        Создание индикатора LTV/CAC.
        
        Args:
            result: Результаты расчета
            
        Returns:
            go.Figure: Индикаторная диаграмма
        """
        ltv_cac = min(result.ltv_cac_ratio, 10)  # Ограничиваем для шкалы
        
        # Определяем цвет зоны
        if ltv_cac >= 5:
            color = FBSVisualizer.COLORS['success']
            status = "ОТЛИЧНО"
        elif ltv_cac >= 3:
            color = FBSVisualizer.COLORS['info']
            status = "ХОРОШО"
        elif ltv_cac >= 1:
            color = FBSVisualizer.COLORS['warning']
            status = "ПРИЕМЛЕМО"
        else:
            color = FBSVisualizer.COLORS['danger']
            status = "ПЛОХО"
        
        fig = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=ltv_cac,
            number=dict(
                suffix="x",
                font=dict(size=40, color=color)
            ),
            delta=dict(
                reference=3,
                increasing=dict(color=FBSVisualizer.COLORS['success']),
                decreasing=dict(color=FBSVisualizer.COLORS['danger'])
            ),
            title=dict(
                text=f"<b>LTV / CAC Ratio</b><br><span style='color:{color}'>{status}</span>",
                font=dict(size=16)
            ),
            gauge=dict(
                axis=dict(
                    range=[0, 10],
                    tickwidth=1,
                    tickcolor="darkblue"
                ),
                bar=dict(color=color),
                bgcolor="white",
                borderwidth=2,
                bordercolor="gray",
                steps=[
                    dict(range=[0, 1], color=FBSVisualizer.COLORS['danger']),
                    dict(range=[1, 3], color=FBSVisualizer.COLORS['warning']),
                    dict(range=[3, 5], color=FBSVisualizer.COLORS['info']),
                    dict(range=[5, 10], color=FBSVisualizer.COLORS['success'])
                ],
                threshold=dict(
                    line=dict(color="red", width=4),
                    thickness=0.75,
                    value=3
                )
            )
        ))
        
        fig.update_layout(
            height=350,
            margin=dict(t=50, b=20, l=20, r=20),
            paper_bgcolor='rgba(0,0,0,0)'
        )
        
        return fig

# ============================================================================
# БЛОК 9: ЭКСПОРТ В EXCEL С ЖИВЫМИ ФОРМУЛАМИ И ЛИСТОМ ТАРИФОВ
# ============================================================================

class ProfessionalExcelExporter:
    """
    Профессиональный экспорт в Excel с живыми формулами.
    
    Особенности:
    - Отдельный лист "Тарифы МП" с данными из API
    - Формулы на основном листе ссылаются на лист тарифов
    - При обновлении тарифов в Excel пересчет происходит автоматически
    - Поддержка больших объемов данных
    - Условное форматирование и визуализация
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
        
        Структура отчета:
        1. 📋 Тарифы МП - актуальные тарифы из API
        2. 📊 Юнит-экономика FBS - основной расчет с живыми формулами
        3. 🔄 Сравнение моделей - FBS vs FBO vs FBP
        4. 👥 LTV и CAC - метрики клиентской экономики
        5. ⚠️ Скрытые потери FBS - анализ неочевидных расходов
        6. 📈 Дашборд - сводная визуализация
        7. 🔌 API Интеграция - информация об источниках данных
        8. 📖 Инструкция - руководство по использованию
        
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
            
            # Создаем листы в правильном порядке
            logger.info("📊 Создание листа тарифов...")
            ws_tariffs = wb.create_sheet("📋 Тарифы МП", 0)
            self._create_tariffs_sheet(ws_tariffs, calculator, marketplace_name)
            
            logger.info("📊 Создание основного листа...")
            ws_main = wb.create_sheet("📊 Юнит-экономика FBS")
            self._create_main_sheet_with_tariff_links(
                ws_main, results, input_data_list, marketplace_name, calculator
            )
            
            logger.info("📊 Создание листа сравнения моделей...")
            ws_models = wb.create_sheet("🔄 Сравнение моделей")
            self._create_models_comparison_sheet(ws_models, results, marketplace_name)
            
            logger.info("📊 Создание листа LTV/CAC...")
            ws_ltv = wb.create_sheet("👥 LTV и CAC")
            self._create_ltv_cac_sheet(ws_ltv, results, input_data_list, marketplace_name)
            
            logger.info("📊 Создание листа скрытых потерь...")
            ws_hidden = wb.create_sheet("⚠️ Скрытые потери FBS")
            self._create_hidden_losses_sheet(ws_hidden, results, input_data_list, marketplace_name)
            
            logger.info("📊 Создание дашборда...")
            ws_dashboard = wb.create_sheet("📈 Дашборд")
            self._create_dashboard_sheet(ws_dashboard, results, marketplace_name)
            
            logger.info("📊 Создание листа API...")
            ws_api = wb.create_sheet("🔌 API Интеграция")
            self._create_api_info_sheet(ws_api, calculator, marketplace_name)
            
            logger.info("📊 Создание инструкции...")
            ws_instructions = wb.create_sheet("📖 Инструкция")
            self._create_instructions_sheet(ws_instructions)
            
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
        
        Args:
            ws: Рабочий лист
            calculator: Калькулятор с загруженными тарифами
            marketplace_name: Название маркетплейса
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
        
        Args:
            ws: Рабочий лист
            results: Результаты расчета
            input_data_list: Входные данные
            marketplace_name: Название маркетплейса
            calculator: Калькулятор
        """
        # Название листа тарифов для формул
        tariff_sheet = "'📋 Тарифы МП'"
        
        # Заголовок
        ws.merge_cells('A1:AV1')
        title_cell = ws.cell(row=1, column=1, value=(
            f"🚀 Юнит-экономика FBS — {marketplace_name} — "
            f"{datetime.now().strftime('%d.%m.%Y %H:%M')}"
        ))
        title_cell.font = self.title_font
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 45
        
        # Подзаголовок с легендой
        ws.merge_cells('A2:AV2')
        ws.cell(row=2, column=1, value=(
            "🟡 Вводные данные (редактируемые) | 🟣 Тарифы из API (лист '📋 Тарифы МП') | "
            "🟢 Расчетные формулы | 🔵 Итоговые показатели | "
            "📎 Формулы автоматически подтягивают данные из листа тарифов"
        )).font = Font(size=9, italic=True, color="666666")
        
        # Определяем заголовки колонок
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
            
            # Расходы (W-AH)
            ("Комиссия МП, ₽", 14), ("First Mile, ₽", 13),
            ("Last Mile, ₽", 13), ("Pick & Pack, ₽", 14),
            ("Упаковка расч., ₽", 14), ("Эквайринг, ₽", 13),
            ("Возвраты, ₽", 12), ("Штрафы, ₽", 12),
            ("Маркетинг расч., ₽", 15), ("Складские, ₽", 12),
            ("Налог, ₽", 10),
            
            # Итоговые показатели (AI-AN)
            ("ИТОГО расходов, ₽", 16), ("ЧИСТАЯ ПРИБЫЛЬ, ₽", 16),
            ("МАРЖА, %", 10), ("ROI, %", 10),
            ("Мин. цена, ₽", 12), ("Макс. скидка, %", 13),
            
            # FBS метрики (AO-AT)
            ("Точка безубыт., км", 15), ("LTV, ₽", 12),
            ("CAC, ₽", 12), ("LTV/CAC", 10),
            ("Запас прочности, ₽", 15), ("Рек. модель", 12),
            
            # Сравнение моделей (AU-AV)
            ("Прибыль FBO, ₽", 15), ("Прибыль FBP, ₽", 15)
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
            
            # === РАСХОДЫ С ССЫЛКАМИ НА ЛИСТ ТАРИФОВ ===
            
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
            
            # Максимальная скидка (колонка 39 = AM)
            formula = f"=IF(D{row_idx}>0, ((D{row_idx}-AL{row_idx})/D{row_idx})*100, 0)"
            cell = ws.cell(row=row_idx, column=39, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '0.00"%"'
            
            # Точка безубыточности (колонка 40 = AN)
            formula = f"=IF(X{row_idx}>0, AI{row_idx}/(K{row_idx}*2/L{row_idx}), 999999)"
            cell = ws.cell(row=row_idx, column=40, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.0'
            
            # LTV (колонка 41 = AO)
            formula = f"=D{row_idx}*{input_data.avg_purchases_per_year}*{input_data.customer_retention_rate}/(1+{input_data.discount_rate})"
            cell = ws.cell(row=row_idx, column=41, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # CAC (колонка 42 = AP)
            formula = f"=(AE{row_idx}+AD{row_idx}+X{row_idx})/0.3"
            cell = ws.cell(row=row_idx, column=42, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # LTV/CAC (колонка 43 = AQ)
            formula = f"=IF(AP{row_idx}>0, AO{row_idx}/AP{row_idx}, 999)"
            cell = ws.cell(row=row_idx, column=43, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '0.0'
            
            # Запас прочности (колонка 44 = AR)
            formula = f"=D{row_idx}-AL{row_idx}"
            cell = ws.cell(row=row_idx, column=44, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Рекомендуемая модель (колонка 45 = AS)
            formula = (
                f'=IF(AND(AI{row_idx}>=AU{row_idx}, AI{row_idx}>=AV{row_idx}), "FBS", '
                f'IF(AU{row_idx}>=AV{row_idx}, "FBO", "FBP"))'
            )
            cell = ws.cell(row=row_idx, column=45, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.font = Font(bold=True)
            
            # Прибыль FBO (колонка 46 = AT)
            formula = (
                f"=D{row_idx}-(E{row_idx}+W{row_idx}+Y{row_idx}*{tariff_sheet}!K{tariff_row}+"
                f"AB{row_idx}+AC{row_idx}+AG{row_idx}+AE{row_idx}+AA{row_idx}+10)"
            )
            cell = ws.cell(row=row_idx, column=46, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Прибыль FBP (колонка 47 = AU)
            formula = (
                f"=D{row_idx}-(E{row_idx}+W{row_idx}+Y{row_idx}*{tariff_sheet}!L{tariff_row}+"
                f"AB{row_idx}+AC{row_idx}+AG{row_idx}+AE{row_idx}+AA{row_idx}+5)"
            )
            cell = ws.cell(row=row_idx, column=47, value=formula)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
        
        # Условное форматирование для прибыли
        last_data_row = len(results) + 4
        if last_data_row >= 5:
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
    
    def _create_models_comparison_sheet(self, ws, results, marketplace_name):
        """Создание листа сравнения моделей FBS/FBO/FBP"""
        ws.merge_cells('A1:I1')
        ws.cell(row=1, column=1, value=(
            f"🔄 Сравнение моделей фулфилмента — {marketplace_name}"
        )).font = self.title_font
        
        headers = [
            ("Артикул", 18), ("Наименование", 30),
            ("Прибыль FBS, ₽", 16), ("Прибыль FBO, ₽", 16),
            ("Прибыль FBP, ₽", 16),
            ("Разница FBS-FBO, ₽", 18), ("Разница FBS-FBP, ₽", 18),
            ("Лучшая модель", 14), ("Рекомендация", 40)
        ]
        
        for col_idx, (header_text, width) in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header_text)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.thin_border
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        
        for i, result in enumerate(results, 4):
            ws.cell(row=i, column=1, value=result.artikul).border = self.thin_border
            ws.cell(row=i, column=2, value=result.product_name).border = self.thin_border
            ws.cell(row=i, column=3, value=result.gross_profit).border = self.thin_border
            ws.cell(row=i, column=3).number_format = '#,##0.00'
            ws.cell(row=i, column=4, value=result.fbo_profit).border = self.thin_border
            ws.cell(row=i, column=4).number_format = '#,##0.00'
            ws.cell(row=i, column=5, value=result.fbp_profit).border = self.thin_border
            ws.cell(row=i, column=5).number_format = '#,##0.00'
            
            # Разницы
            diff_fbo = result.gross_profit - result.fbo_profit
            diff_fbp = result.gross_profit - result.fbp_profit
            
            cell_fbo = ws.cell(row=i, column=6, value=diff_fbo)
            cell_fbo.border = self.thin_border
            cell_fbo.number_format = '#,##0.00'
            if diff_fbo > 0:
                cell_fbo.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            else:
                cell_fbo.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            
            cell_fbp = ws.cell(row=i, column=7, value=diff_fbp)
            cell_fbp.border = self.thin_border
            cell_fbp.number_format = '#,##0.00'
            if diff_fbp > 0:
                cell_fbp.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            else:
                cell_fbp.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            
            # Лучшая модель
            model_cell = ws.cell(row=i, column=8, value=result.recommended_model)
            model_cell.border = self.thin_border
            model_cell.font = Font(bold=True)
            
            if result.recommended_model == "FBS":
                model_cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            elif result.recommended_model == "FBO":
                model_cell.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
            else:
                model_cell.fill = PatternFill(start_color="E0E0FF", end_color="E0E0FF", fill_type="solid")
            
            # Рекомендация
            if result.recommended_model == "FBS":
                recommendation = "Оставить на FBS. Текущая модель оптимальна."
            elif result.recommended_model == "FBO":
                recommendation = f"Перейти на FBO. Экономия: {abs(diff_fbo):,.0f} ₽ за счет снижения логистики."
            else:
                recommendation = f"Перейти на FBP. Экономия: {abs(diff_fbp):,.0f} ₽ за счет оптимизации."
            
            ws.cell(row=i, column=9, value=recommendation).border = self.thin_border
    
    def _create_ltv_cac_sheet(self, ws, results, input_data_list, marketplace_name):
        """Создание листа с метриками LTV и CAC"""
        ws.merge_cells('A1:J1')
        ws.cell(row=1, column=1, value=f"👥 Метрики LTV и CAC — {marketplace_name}").font = self.title_font
        
        headers = [
            ("Артикул", 18), ("Наименование", 30),
            ("Средний чек, ₽", 15), ("Покупок в год", 13),
            ("Коэфф. удержания", 15), ("LTV, ₽", 13),
            ("CAC, ₽", 13), ("LTV/CAC", 10),
            ("Оценка", 15), ("Рекомендация", 45)
        ]
        
        for col_idx, (header_text, width) in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header_text)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.thin_border
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        
        for i, (result, input_data) in enumerate(zip(results, input_data_list), 4):
            ws.cell(row=i, column=1, value=result.artikul).border = self.thin_border
            ws.cell(row=i, column=2, value=result.product_name).border = self.thin_border
            ws.cell(row=i, column=3, value=result.selling_price).border = self.thin_border
            ws.cell(row=i, column=4, value=input_data.avg_purchases_per_year).border = self.thin_border
            ws.cell(row=i, column=5, value=input_data.customer_retention_rate).border = self.thin_border
            ws.cell(row=i, column=6, value=result.ltv).border = self.thin_border
            ws.cell(row=i, column=6).number_format = '#,##0.00'
            ws.cell(row=i, column=7, value=result.cac).border = self.thin_border
            ws.cell(row=i, column=7).number_format = '#,##0.00'
            
            # LTV/CAC
            ltv_cac_cell = ws.cell(row=i, column=8, value=result.ltv_cac_ratio)
            ltv_cac_cell.border = self.thin_border
            ltv_cac_cell.number_format = '0.0'
            
            if result.ltv_cac_ratio >= 5:
                ltv_cac_cell.fill = PatternFill(start_color="006100", end_color="006100", fill_type="solid")
                ltv_cac_cell.font = Font(color="FFFFFF", bold=True)
                assessment = "✅ ОТЛИЧНО"
                recommendation = "Масштабируйте рекламный бюджет. Отличная юнит-экономика."
            elif result.ltv_cac_ratio >= 3:
                ltv_cac_cell.fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
                assessment = "✅ Хорошо"
                recommendation = "Продолжайте оптимизацию. Есть потенциал для роста."
            elif result.ltv_cac_ratio >= 1:
                ltv_cac_cell.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                assessment = "⚠️ Приемлемо"
                recommendation = "Работайте над удержанием клиентов и снижением CAC."
            else:
                ltv_cac_cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                assessment = "❌ Плохо"
                recommendation = "Срочно пересмотрите стратегию! CAC превышает LTV."
            
            ws.cell(row=i, column=9, value=assessment).border = self.thin_border
            ws.cell(row=i, column=10, value=recommendation).border = self.thin_border
    
    def _create_hidden_losses_sheet(self, ws, results, input_data_list, marketplace_name):
        """Создание листа анализа скрытых потерь FBS"""
        ws.merge_cells('A1:G1')
        ws.cell(row=1, column=1, value=f"⚠️ Анализ скрытых потерь FBS — {marketplace_name}").font = self.title_font
        
        headers = [
            ("Статья скрытых потерь", 35),
            ("Формула расчета", 50),
            ("Среднее значение", 20),
            ("Максимальное", 18),
            ("Уровень риска", 15),
            ("Влияние на прибыль", 18),
            ("Рекомендация", 50)
        ]
        
        for col_idx, (header_text, width) in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header_text)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.thin_border
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        
        if results:
            # Расчет статистики
            avg_first_mile = np.mean([r.first_mile_cost for r in results])
            max_first_mile = max([r.first_mile_cost for r in results])
            
            avg_penalty = np.mean([r.penalty_cost for r in results])
            max_penalty = max([r.penalty_cost for r in results])
            
            avg_pick_pack = np.mean([r.pick_pack_cost for r in results])
            max_pick_pack = max([r.pick_pack_cost for r in results])
            
            avg_total_logistics = np.mean([r.first_mile_cost + r.last_mile_cost for r in results])
            
            avg_price = np.mean([r.selling_price for r in results])
            
            hidden_losses = [
                [
                    "Двойная логистика (First + Last Mile)",
                    "First Mile Cost + Last Mile Cost",
                    f"{avg_total_logistics:,.2f} ₽",
                    f"{max([r.first_mile_cost + r.last_mile_cost for r in results]):,.2f} ₽",
                    "КРИТИЧЕСКИЙ" if avg_total_logistics / avg_price > 0.25 else "ВЫСОКИЙ",
                    f"{avg_total_logistics / avg_price * 100:.1f}% от цены",
                    "Оптимизируйте First Mile: увеличьте загрузку паллет, пересмотрите маршруты"
                ],
                [
                    "Штрафы за просрочку передачи заказа",
                    "Цена × Ставка штрафа × Вероятность просрочки",
                    f"{avg_penalty:,.2f} ₽",
                    f"{max_penalty:,.2f} ₽",
                    "ВЫСОКИЙ" if avg_penalty / avg_price > 0.02 else "СРЕДНИЙ",
                    f"{avg_penalty / avg_price * 100:.1f}% от цены",
                    "Внедрите ночную смену или ускорьте обработку заказов"
                ],
                [
                    "Стоимость обработки заказа (Pick & Pack)",
                    "(Время сборки / 60) × Ставка оператора",
                    f"{avg_pick_pack:,.2f} ₽",
                    f"{max_pick_pack:,.2f} ₽",
                    "СРЕДНИЙ",
                    f"{avg_pick_pack / avg_price * 100:.1f}% от цены",
                    "Оптимизируйте складские процессы, обучите персонал"
                ],
                [
                    "Износ упаковки и повреждения при транспортировке",
                    "Стоимость упаковки + Риск повреждения × Цена",
                    f"{np.mean([r.packaging_cost for r in results]):,.2f} ₽",
                    f"{max([r.packaging_cost for r in results]):,.2f} ₽",
                    "НИЗКИЙ",
                    f"{np.mean([r.packaging_cost for r in results]) / avg_price * 100:.1f}% от цены",
                    "Используйте качественную упаковку для снижения повреждений"
                ],
                [
                    "Складские расходы на хранение запаса",
                    "Аренда склада / Оборачиваемость",
                    f"{np.mean([r.warehouse_cost for r in results]):,.2f} ₽",
                    f"{max([r.warehouse_cost for r in results]):,.2f} ₽",
                    "СРЕДНИЙ",
                    f"{np.mean([r.warehouse_cost for r in results]) / avg_price * 100:.1f}% от цены",
                    "Оптимизируйте глубину запаса, используйте ABC-анализ"
                ]
            ]
            
            for row_idx, row_data in enumerate(hidden_losses, 4):
                for col_idx, value in enumerate(row_data, 1):
                    cell = ws.cell(row=row_idx, column=col_idx, value=value)
                    cell.border = self.thin_border
                    
                    # Цветовое кодирование уровня риска
                    if col_idx == 5:
                        if "КРИТИЧЕСКИЙ" in str(value):
                            cell.fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
                        elif "ВЫСОКИЙ" in str(value):
                            cell.fill = PatternFill(start_color="FFF3CD", end_color="FFF3CD", fill_type="solid")
                        elif "СРЕДНИЙ" in str(value):
                            cell.fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                        else:
                            cell.fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
    
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
        
        metrics = [
            ("Всего товаров", f"{total_items}", "A4"),
            ("Прибыльных", f"{profitable_items} ({profitable_items/total_items*100:.0f}%)", "C4"),
            ("Убыточных", f"{unprofitable_items} ({unprofitable_items/total_items*100:.0f}%)", "E4"),
            ("Общая прибыль", f"{total_profit:,.0f} ₽", "G4"),
            ("Общая выручка", f"{total_revenue:,.0f} ₽", "A7"),
            ("Общие расходы", f"{total_costs:,.0f} ₽", "C7"),
            ("Средняя маржа", f"{avg_margin:.1f}%", "E7"),
            ("Средний ROI", f"{avg_roi:.1f}%", "G7")
        ]
        
        for title, value, cell_ref in metrics:
            row = int(cell_ref[1:])
            col = ord(cell_ref[0]) - ord('A') + 1
            
            title_cell = ws.cell(row=row, column=col, value=title)
            title_cell.font = Font(bold=True, size=11, color="666666")
            
            value_cell = ws.cell(row=row+1, column=col, value=value)
            value_cell.font = Font(bold=True, size=16, color="1a1a2e")
    
    def _create_api_info_sheet(self, ws, calculator, marketplace_name):
        """Создание листа с информацией об API интеграции"""
        ws.merge_cells('A1:E1')
        ws.cell(row=1, column=1, value=f"🔌 Информация об API интеграции — {marketplace_name}").font = self.title_font
        
        api_manager = calculator.api_manager
        
        api_info = [
            ["Сервис", "Эндпоинт", "Статус", "Последняя проверка", "Действие"],
            [
                "Ozon API",
                "https://api.ozon.ru/v1",
                "✅ Подключен" if api_manager.has_api_key('ozon') else "⚠️ Не настроен",
                datetime.now().strftime('%d.%m.%Y %H:%M'),
                "Настроить API ключи в приложении"
            ],
            [
                "Wildberries API",
                "https://suppliers-api.wildberries.ru",
                "✅ Подключен" if api_manager.has_api_key('wildberries') else "⚠️ Не настроен",
                datetime.now().strftime('%d.%m.%Y %H:%M'),
                "Настроить API ключи в приложении"
            ],
            [
                "Яндекс Маркет API",
                "https://api.partner.market.yandex.ru",
                "✅ Подключен" if api_manager.has_api_key('yandex_market') else "⚠️ Не настроен",
                datetime.now().strftime('%d.%m.%Y %H:%M'),
                "Настроить API ключи в приложении"
            ],
            [
                "DeepSeek AI",
                "https://api.deepseek.com",
                "✅ Подключен" if api_manager.has_api_key('deepseek') else "⚠️ Не настроен",
                datetime.now().strftime('%d.%m.%Y %H:%M'),
                "Используется для AI-обогащения тарифов"
            ],
            [
                "Кэш тарифов",
                f"TTL: 1 час",
                "✅ Активен",
                datetime.now().strftime('%d.%m.%Y %H:%M'),
                "Очистить кэш в приложении"
            ]
        ]
        
        for row_idx, row_data in enumerate(api_info, 3):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self.thin_border
                if row_idx == 3:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="E8E8E8", end_color="E8E8E8", fill_type="solid")
    
    def _create_instructions_sheet(self, ws):
        """Создание листа с инструкцией по использованию"""
        ws.merge_cells('A1:D1')
        ws.cell(row=1, column=1, value="📖 Инструкция по использованию отчета").font = self.title_font
        
        instructions = [
            ["Раздел отчета", "Описание", "Как использовать", "Важно"],
            [
                "📋 Тарифы МП",
                "Актуальные тарифы маркетплейса, загруженные через API",
                "Проверяйте актуальность. При необходимости обновите через приложение.",
                "Данные на этом листе используются в формулах других листов"
            ],
            [
                "📊 Юнит-экономика FBS",
                "Основной расчет прибыльности товаров",
                "Изменяйте желтые ячейки (вводные данные). Зеленые ячейки пересчитываются автоматически.",
                "Формулы содержат ссылки на лист '📋 Тарифы МП'"
            ],
            [
                "🔄 Сравнение моделей",
                "Сравнение FBS, FBO и FBP для каждого товара",
                "Используйте для принятия решения о смене модели фулфилмента.",
                "Зеленый цвет — текущая модель оптимальна, желтый — стоит пересмотреть"
            ],
            [
                "👥 LTV и CAC",
                "Метрики клиентской экономики",
                "LTV/CAC > 3 — отлично, > 1 — приемлемо, < 1 — требуется оптимизация.",
                "Низкий LTV/CAC означает убыточность на уровне клиента"
            ],
            [
                "⚠️ Скрытые потери FBS",
                "Анализ неочевидных расходов",
                "Обратите внимание на статьи с высоким уровнем риска.",
                "Штрафы за просрочку и Pick & Pack часто недооцениваются"
            ],
            [
                "📈 Дашборд",
                "Сводная статистика",
                "Быстрая оценка общего состояния портфеля товаров.",
                "Обновляется автоматически при изменении данных"
            ],
            [
                "🔌 API Интеграция",
                "Статус подключений к API",
                "Проверьте, что все API ключи настроены для получения актуальных тарифов.",
                "Без API используются дефолтные значения, которые могут быть неточными"
            ]
        ]
        
        for row_idx, row_data in enumerate(instructions, 3):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self.thin_border
                if row_idx == 3:
                    cell.font = Font(bold=True, color="FFFFFF")
                    cell.fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
                elif col_idx == 1:
                    cell.font = Font(bold=True)
        
        ws.column_dimensions['A'].width = 22
        ws.column_dimensions['B'].width = 35
        ws.column_dimensions['C'].width = 40
        ws.column_dimensions['D'].width = 40

# ============================================================================
# БЛОК 10: ИНТЕРФЕЙС ПОЛЬЗОВАТЕЛЯ (STREAMLIT)
# ============================================================================

def init_session_state():
    """Инициализация всех состояний сессии Streamlit"""
    
    if 'api_manager' not in st.session_state:
        st.session_state.api_manager = MarketplaceAPIManager()
    
    if 'calculator' not in st.session_state:
        st.session_state.calculator = FBSUnitEconomicsCalculator(
            api_manager=st.session_state.api_manager
        )
    
    if 'visualizer' not in st.session_state:
        st.session_state.visualizer = FBSVisualizer()
    
    if 'secure_data' not in st.session_state:
        st.session_state.secure_data = SecureDataManager()
    
    if 'results' not in st.session_state:
        st.session_state.results = []
    
    if 'input_data_list' not in st.session_state:
        st.session_state.input_data_list = []
    
    if 'exporter' not in st.session_state:
        try:
            st.session_state.exporter = ProfessionalExcelExporter()
        except ImportError:
            st.session_state.exporter = None
    
    if 'marketplace' not in st.session_state:
        st.session_state.marketplace = "Ozon"
    
    if 'tax_system' not in st.session_state:
        st.session_state.tax_system = "УСН 6% (доходы)"
    
    if 'current_section' not in st.session_state:
        st.session_state.current_section = 'main'
    
    if 'show_api_settings' not in st.session_state:
        st.session_state.show_api_settings = False

def render_sidebar():
    """Отрисовка боковой панели навигации"""
    
    with st.sidebar:
        # Заголовок
        st.markdown("""
        <div style='text-align: center; padding: 20px 15px; background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460); border-radius: 12px; margin-bottom: 25px;'>
            <h1 style='color: white; margin: 0; font-size: 1.5em;'>🚀 FBS PRO</h1>
            <p style='color: #a8a8d0; margin: 8px 0 0 0; font-size: 0.9em;'>Юнит-экономика 2026</p>
            <p style='color: #6666aa; margin: 5px 0 0 0; font-size: 0.7em;'>v{APP_VERSION} | API Ready</p>
        </div>
        """.format(APP_VERSION=APP_VERSION), unsafe_allow_html=True)
        
        # Навигация
        st.markdown("### 🧭 Навигация")
        
        sections = {
            "🏠 Главная": "main",
            "🧮 Калькулятор FBS": "calculator",
            "📋 Тарифы маркетплейсов": "tariffs",
            "📈 Дашборд": "dashboard",
            "📥 Экспорт в Excel": "export",
            "⚙️ Настройки": "settings"
        }
        
        selected_section = st.radio(
            "Выберите раздел:",
            list(sections.keys()),
            label_visibility="collapsed"
        )
        
        st.session_state.current_section = sections[selected_section]
        
        # Статус системы
        st.markdown("---")
        st.markdown("### 📊 Статус системы")
        
        calculator = st.session_state.calculator
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown(f"**🏪 МП:** {st.session_state.marketplace}")
        with col2:
            st.markdown(f"**💰 Налог:** {st.session_state.tax_system.split()[0]}")
        
        # Статус тарифов
        if calculator.tariffs_source == 'api':
            st.success("🔌 Тарифы: API")
        elif calculator.tariffs_source == 'deepseek':
            st.info("🤖 Тарифы: DeepSeek AI")
        else:
            st.warning("⚠️ Тарифы: Дефолтные")
        
        # Статус API ключей
        api_manager = st.session_state.api_manager
        api_status = []
        
        if api_manager.has_api_key('ozon') and api_manager.has_api_key('ozon_client_id'):
            api_status.append("✅ Ozon")
        else:
            api_status.append("⚪ Ozon")
        
        if api_manager.has_api_key('wildberries'):
            api_status.append("✅ WB")
        else:
            api_status.append("⚪ WB")
        
        if api_manager.has_api_key('yandex_market'):
            api_status.append("✅ YM")
        else:
            api_status.append("⚪ YM")
        
        if api_manager.has_api_key('deepseek'):
            api_status.append("✅ AI")
        else:
            api_status.append("⚪ AI")
        
        st.markdown("**🔑 API ключи:**")
        st.markdown(" | ".join(api_status))
        
        # Результаты расчетов
        if st.session_state.results:
            st.success(f"✅ Рассчитано: {len(st.session_state.results)} товаров")
        else:
            st.info("ℹ️ Расчеты не выполнялись")
        
        # Быстрые действия
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
        
        # Информация о приложении
        st.markdown("---")
        st.markdown(f"""
        <div style='text-align: center; color: #888; font-size: 0.75em;'>
            <p>🚀 FBS Unit Economics PRO</p>
            <p>Версия {APP_VERSION}</p>
            <p>© 2026 | API Integration</p>
        </div>
        """, unsafe_allow_html=True)

def show_main_page():
    """Главная страница приложения"""
    
    # Приветственный баннер
    st.markdown("""
    <div style='text-align: center; padding: 50px 30px; background: linear-gradient(135deg, #0f0c29, #302b63, #24243e); border-radius: 20px; margin-bottom: 35px;'>
        <h1 style='color: white; font-size: 3em; margin: 0;'>🚀 FBS Юнит-экономика PRO</h1>
        <p style='color: #a8a8d0; font-size: 1.3em; margin: 20px 0;'>
            Профессиональный расчет юнит-экономики для FBS-модели с автообновлением тарифов через API
        </p>
        <p style='color: #6666aa; font-size: 1em; margin: 10px 0;'>
            Ozon • Wildberries • Яндекс Маркет | API Integration | DeepSeek AI | Excel с живыми формулами
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Ключевые возможности
    st.markdown("### 🎯 Ключевые возможности")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #0984e3, #6c5ce7); padding: 25px; border-radius: 15px; color: white; height: 100%;'>
            <h3 style='margin-top: 0;'>🔌 API Интеграция</h3>
            <p>Автоматическая загрузка актуальных тарифов через API маркетплейсов</p>
            <ul>
                <li>Ozon API</li>
                <li>Wildberries API</li>
                <li>Яндекс Маркет API</li>
            </ul>
            <p><strong>Всегда актуальные данные для расчетов</strong></p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #00b894, #00cec9); padding: 25px; border-radius: 15px; color: white; height: 100%;'>
            <h3 style='margin-top: 0;'>🤖 AI Обогащение</h3>
            <p>DeepSeek AI для получения тарифов когда прямые API недоступны</p>
            <ul>
                <li>Автоматический фолбэк</li>
                <li>Верификация данных</li>
                <li>Анализ конкурентов</li>
            </ul>
            <p><strong>Умное резервирование источников данных</strong></p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #e17055, #d63031); padding: 25px; border-radius: 15px; color: white; height: 100%;'>
            <h3 style='margin-top: 0;'>📊 Excel с формулами</h3>
            <p>Профессиональные отчеты с живыми формулами и листом тарифов</p>
            <ul>
                <li>Ссылки на лист тарифов</li>
                <li>Автопересчет при обновлении</li>
                <li>8 листов с аналитикой</li>
            </ul>
            <p><strong>Обновите тарифы — всё пересчитается</strong></p>
        </div>
        """, unsafe_allow_html=True)
    
    # Быстрый старт
    st.markdown("---")
    st.markdown("### 🚀 Быстрый старт")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.info("""
        **Шаг 1: Настройте API ключи**
        
        Перейдите в раздел ⚙️ Настройки и добавьте API ключи маркетплейсов для автоматической загрузки тарифов.
        """)
    
    with col2:
        st.info("""
        **Шаг 2: Проверьте тарифы**
        
        В разделе 📋 Тарифы маркетплейсов проверьте актуальность загруженных данных.
        """)
    
    with col3:
        st.info("""
        **Шаг 3: Выполните расчет**
        
        Используйте 🧮 Калькулятор FBS для расчета юнит-экономики и экспортируйте результаты в Excel.
        """)
    
    # Статистика последних расчетов
    if st.session_state.results:
        st.markdown("---")
        st.markdown("### 📈 Статистика последних расчетов")
        
        results = st.session_state.results
        
        col1, col2, col3, col4, col5 = st.columns(5)
        
        with col1:
            st.metric("Товаров", len(results))
        with col2:
            avg_margin = np.mean([r.margin_percent for r in results])
            st.metric("Средняя маржа", f"{avg_margin:.1f}%")
        with col3:
            profitable = len([r for r in results if r.gross_profit > 0])
            st.metric("Прибыльных", f"{profitable}")
        with col4:
            total_profit = sum(r.gross_profit for r in results)
            st.metric("Общая прибыль", f"{total_profit:,.0f} ₽")
        with col5:
            fbo_better = len([r for r in results if r.fbo_profit > r.gross_profit])
            st.metric("FBO выгоднее", f"{fbo_better}")

def show_calculator_page():
    """Страница калькулятора FBS"""
    
    st.markdown("## 🧮 Калькулятор FBS юнит-экономики")
    
    # Информационное сообщение
    tariff_source_badge = ""
    if st.session_state.calculator.tariffs_source == 'api':
        tariff_source_badge = "🟢 Тарифы загружены через API"
    elif st.session_state.calculator.tariffs_source == 'deepseek':
        tariff_source_badge = "🔵 Тарифы получены через DeepSeek AI"
    else:
        tariff_source_badge = "🟡 Используются дефолтные тарифы"
    
    st.info(f"""
    **🎯 Этот калькулятор использует актуальные тарифы: {tariff_source_badge}**
    
    **Учитывает специфику FBS:**
    - 🚛 **First Mile** — ваша логистика до склада маркетплейса
    - 📦 **Last Mile** — доставка маркетплейса до клиента
    - ⚠️ **Штрафы за просрочку** — с расчетом вероятности
    - 👷 **Pick & Pack** — стоимость обработки заказа на вашем складе
    - 📏 **Точка безубыточности** — максимальное расстояние до склада МП
    - 💰 **Запас прочности** — максимальная скидка для распродаж
    - 👥 **LTV и CAC** — метрики клиентской экономики
    - 🔄 **Сравнение FBS/FBO/FBP** — выбор оптимальной модели
    """)
    
    # Выбор режима расчета
    calc_mode = st.radio(
        "Режим расчета:",
        ["📱 Расчет одного товара", "📊 Массовый расчет из файла"],
        horizontal=True,
        key="calc_mode"
    )
    
    if calc_mode == "📱 Расчет одного товара":
        show_single_calculator()
    else:
        show_batch_calculator()

def show_single_calculator():
    """Калькулятор для одного товара"""
    
    calculator = st.session_state.calculator
    
    # Создаем вкладки для группировки параметров
    tab1, tab2, tab3 = st.tabs(["📦 Товар и цены", "🚚 Логистика FBS", "📊 Маркетинг и клиенты"])
    
    with tab1:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Основные параметры")
            
            artikul = st.text_input("Артикул товара", value="SKU-001", key="single_artikul")
            product_name = st.text_input("Наименование товара", value="Тестовый товар", key="single_name")
            
            # Категория из загруженных тарифов
            available_categories = list(calculator.current_tariffs.keys()) if calculator.current_tariffs else ["default"]
            category = st.selectbox(
                "Категория товара",
                options=available_categories,
                index=0,
                key="single_category",
                help="Выберите категорию — от этого зависят применяемые тарифы"
            )
        
        with col2:
            st.subheader("Финансовые параметры")
            
            selling_price = st.number_input(
                "Цена продажи на маркетплейсе, ₽",
                value=5000.0, step=100.0, min_value=1.0,
                key="single_price",
                help="Розничная цена, по которой товар продается на маркетплейсе"
            )
            
            cogs = st.number_input(
                "Себестоимость закупки (COGS), ₽",
                value=3000.0, step=100.0, min_value=1.0,
                key="single_cogs",
                help="Закупочная цена товара у поставщика"
            )
        
        st.markdown("---")
        st.subheader("📏 Габариты и вес товара")
        
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            weight = st.number_input(
                "Вес брутто, кг",
                value=1.5, step=0.1, min_value=0.01,
                key="single_weight",
                help="Вес товара в упаковке (брутто)"
            )
        
        with col2:
            length = st.number_input(
                "Длина, см",
                value=20, step=1, min_value=0,
                key="single_length"
            )
        
        with col3:
            width = st.number_input(
                "Ширина, см",
                value=15, step=1, min_value=0,
                key="single_width"
            )
        
        with col4:
            height = st.number_input(
                "Высота, см",
                value=10, step=1, min_value=0,
                key="single_height"
            )
    
    with tab2:
        st.subheader("🚛 First Mile — Ваша логистика до склада маркетплейса")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            warehouse_distance = st.number_input(
                "Расстояние до склада МП, км",
                value=50.0, step=1.0, min_value=0.0,
                key="single_distance",
                help="Расстояние от вашего склада до ближайшего склада маркетплейса"
            )
        
        with col2:
            transport_cost_per_km = st.number_input(
                "Стоимость 1 км транспорта, ₽",
                value=20.0, step=1.0, min_value=1.0,
                key="single_km_cost",
                help="Стоимость пробега вашего транспорта за 1 км"
            )
        
        with col3:
            pallet_capacity = st.number_input(
                "Единиц товара на паллете",
                value=100, step=10, min_value=1,
                key="single_pallet",
                help="Сколько единиц товара помещается на одной паллете"
            )
        
        st.markdown("---")
        st.subheader("📦 Обработка заказа (Pick & Pack)")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            pick_pack_time = st.number_input(
                "Время сборки одного заказа, мин",
                value=5.0, step=0.5, min_value=0.5,
                key="single_pick_time",
                help="Среднее время, которое оператор тратит на сборку одного заказа"
            )
        
        with col2:
            operator_rate = st.number_input(
                "Ставка оператора склада, ₽/час",
                value=300.0, step=50.0, min_value=100.0,
                key="single_operator_rate",
                help="Почасовая оплата сотрудника склада"
            )
        
        with col3:
            packaging_cost = st.number_input(
                "Стоимость упаковочных материалов, ₽/шт",
                value=50.0, step=10.0, min_value=0.0,
                key="single_packaging",
                help="Затраты на коробку, наполнитель, скотч и т.д."
            )
        
        st.markdown("---")
        st.subheader("⚠️ Риски и штрафы")
        
        has_night_shift = st.checkbox(
            "Наличие ночной смены на складе",
            value=False,
            key="single_night_shift",
            help="Снижает вероятность штрафа за просрочку с 35% до 5%"
        )
        
        if has_night_shift:
            st.success("✅ Ночная смена снижает риск штрафов за просрочку до 5%")
        else:
            st.warning("⚠️ Без ночной смены ~35% заказов рискуют получить штраф за просрочку")
    
    with tab3:
        st.subheader("📊 Маркетинговые расходы")
        
        marketing_budget = st.number_input(
            "Маркетинговый бюджет на единицу товара, ₽",
            value=100.0, step=10.0, min_value=0.0,
            key="single_marketing",
            help="Средние расходы на рекламу и продвижение в пересчете на одну продажу"
        )
        
        st.markdown("---")
        st.subheader("👥 Параметры для расчета LTV и CAC")
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            avg_purchases = st.number_input(
                "Среднее количество покупок в год",
                value=2.5, step=0.1, min_value=0.1,
                key="single_purchases",
                help="Сколько раз в среднем один клиент покупает за год"
            )
        
        with col2:
            crr = st.number_input(
                "Коэффициент удержания (CRR)",
                value=0.7, step=0.05, min_value=0.0, max_value=1.0,
                key="single_crr",
                help="Доля клиентов, которые совершают повторные покупки"
            )
        
        with col3:
            discount_rate = st.number_input(
                "Ставка дисконтирования",
                value=0.1, step=0.01, min_value=0.0, max_value=1.0,
                key="single_discount",
                help="Годовая ставка дисконтирования для расчета LTV"
            )
    
    # Кнопка расчета
    st.markdown("---")
    
    col1, col2, col3 = st.columns([2, 1, 2])
    
    with col2:
        calculate_button = st.button(
            "🚀 Рассчитать FBS",
            type="primary",
            width="stretch",
            key="single_calc_button"
        )
    
    if calculate_button:
        with st.spinner("Выполняется профессиональный расчет FBS с актуальными тарифами..."):
            # Создание входных данных
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
                transport_cost_per_km=transport_cost_per_km,
                pallet_capacity=pallet_capacity,
                packaging_cost=packaging_cost,
                pick_pack_time_min=pick_pack_time,
                operator_hourly_rate=operator_rate,
                marketing_budget_per_unit=marketing_budget,
                has_night_shift=has_night_shift,
                avg_purchases_per_year=avg_purchases,
                customer_retention_rate=crr,
                discount_rate=discount_rate
            )
            
            # Валидация
            validation_errors = input_data.validate()
            if validation_errors:
                for error in validation_errors:
                    st.warning(f"⚠️ {error}")
            
            # Выполнение расчета
            result = calculator.calculate_unit_economics(input_data)
            
            # Сохранение результатов
            st.session_state.results = [result]
            st.session_state.input_data_list = [input_data]
            
            # Отображение результатов
            show_single_result(result, input_data)

def show_single_result(result: FBSResultData, input_data: FBSInputData):
    """Отображение результатов расчета одного товара"""
    
    st.markdown("---")
    st.markdown("## 📊 Результаты расчета FBS")
    
    # Ключевые KPI
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        profit_color = "normal" if result.gross_profit > 0 else "inverse"
        st.metric(
            "💰 Чистая прибыль",
            f"{result.gross_profit:,.2f} ₽",
            delta=f"{result.margin_percent:.1f}% маржи",
            delta_color=profit_color
        )
    
    with col2:
        st.metric("📦 Общие расходы", f"{result.total_expenses:,.2f} ₽")
    
    with col3:
        st.metric("📈 ROI", f"{result.roi_percent:.1f}%")
    
    with col4:
        st.metric("💡 Рекомендуемая модель", result.recommended_model)
    
    with col5:
        ltv_cac_color = "normal" if result.ltv_cac_ratio >= 3 else ("off" if result.ltv_cac_ratio >= 1 else "inverse")
        st.metric("👥 LTV/CAC", f"{result.ltv_cac_ratio:.1f}x", delta_color=ltv_cac_color)
    
    # Детализация расходов
    st.markdown("### 📋 Детализация расходов FBS")
    
    expenses_data = {
        "Статья расходов": [
            "Себестоимость закупки (COGS)",
            "Комиссия маркетплейса",
            "🚛 First Mile (доставка до МП)",
            "📦 Last Mile (доставка клиенту)",
            "👷 Pick & Pack (обработка заказа)",
            "📦 Упаковочные материалы",
            "💳 Эквайринг",
            "↩️ Возвраты",
            "⚠️ Штрафы за просрочку",
            "📊 Маркетинговые расходы",
            "🏭 Складские расходы",
            "💰 Налоги",
            "**ИТОГО РАСХОДОВ**",
            "**ЧИСТАЯ ПРИБЫЛЬ**"
        ],
        "Сумма, ₽": [
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
            result.gross_profit
        ],
        "% от цены": [
            f"{input_data.cogs / result.selling_price * 100:.1f}%",
            f"{result.commission / result.selling_price * 100:.1f}%",
            f"{result.first_mile_cost / result.selling_price * 100:.1f}%",
            f"{result.last_mile_cost / result.selling_price * 100:.1f}%",
            f"{result.pick_pack_cost / result.selling_price * 100:.1f}%",
            f"{result.packaging_cost / result.selling_price * 100:.1f}%",
            f"{result.acquiring_cost / result.selling_price * 100:.1f}%",
            f"{result.return_cost / result.selling_price * 100:.1f}%",
            f"{result.penalty_cost / result.selling_price * 100:.1f}%",
            f"{result.marketing_cost / result.selling_price * 100:.1f}%",
            f"{result.warehouse_cost / result.selling_price * 100:.1f}%",
            f"{result.tax_cost / result.selling_price * 100:.1f}%",
            f"**{result.total_expenses / result.selling_price * 100:.1f}%**",
            f"**{abs(result.gross_profit) / result.selling_price * 100:.1f}%**"
        ]
    }
    
    df_expenses = pd.DataFrame(expenses_data)
    
    # Стилизация таблицы
    def style_expenses(row):
        if 'ИТОГО' in str(row['Статья расходов']):
            return ['background-color: #FFF3CD; font-weight: bold'] * len(row)
        elif 'ПРИБЫЛЬ' in str(row['Статья расходов']):
            return ['background-color: #C6EFCE; font-weight: bold'] * len(row)
        return [''] * len(row)
    
    st.dataframe(
        df_expenses.style.apply(style_expenses, axis=1),
        width="stretch",
        height=500
    )
    
    # FBS специфические метрики
    st.markdown("### ⚠️ Специфические метрики FBS")
    
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Вероятность просрочки", f"{result.penalty_probability:.0%}")
        st.metric("Штрафы за просрочку", f"{result.penalty_cost:.2f} ₽")
    
    with col2:
        st.metric("Точка безубыточности", f"{result.break_even_distance_km:.1f} км")
        if result.break_even_distance_km < 25:
            st.error("⚠️ Критически близко!")
        elif result.break_even_distance_km < 50:
            st.warning("⚠️ Зона риска")
        else:
            st.success("✅ Хороший запас")
    
    with col3:
        st.metric("Максимальная скидка", f"{result.max_discount_percent:.1f}%")
        st.metric("Запас прочности", f"{result.safety_margin_price:.2f} ₽")
    
    with col4:
        st.metric("LTV", f"{result.ltv:,.2f} ₽")
        st.metric("CAC", f"{result.cac:,.2f} ₽")
    
    # Сравнение моделей
    st.markdown("### 🔄 Сравнение моделей фулфилмента")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        if result.recommended_model == "FBS":
            st.success(f"**FBS: {result.gross_profit:,.2f} ₽** ✅ Текущая")
        else:
            st.info(f"**FBS: {result.gross_profit:,.2f} ₽**")
    
    with col2:
        if result.recommended_model == "FBO":
            st.success(f"**FBO: {result.fbo_profit:,.2f} ₽** ✅ Рекомендуется")
        else:
            st.info(f"**FBO: {result.fbo_profit:,.2f} ₽**")
    
    with col3:
        if result.recommended_model == "FBP":
            st.success(f"**FBP: {result.fbp_profit:,.2f} ₽** ✅ Рекомендуется")
        else:
            st.info(f"**FBP: {result.fbp_profit:,.2f} ₽**")
    
    # Визуализация
    st.markdown("---")
    st.markdown("## 📊 Визуализация")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig_pie = st.session_state.visualizer.create_cost_breakdown_pie(result)
        st.plotly_chart(fig_pie, width="stretch")
    
    with col2:
        fig_waterfall = st.session_state.visualizer.create_waterfall_chart(result)
        st.plotly_chart(fig_waterfall, width="stretch")
    
    # Рекомендации
    st.markdown("---")
    st.markdown("## 💡 Рекомендации по оптимизации")
    
    recommendations = []
    
    if result.gross_profit <= 0:
        recommendations.append("❌ **Товар убыточен!** Срочно пересмотрите цену или найдите поставщика с более низкой себестоимостью.")
    elif result.margin_percent < 10:
        recommendations.append("⚠️ **Низкая маржинальность.** Рассмотрите возможность повышения цены или снижения расходов.")
    
    if result.first_mile_cost > result.selling_price * 0.15:
        recommendations.append("⚠️ **Высокая стоимость First Mile!** Оптимизируйте логистику или рассмотрите переход на FBO.")
    
    if result.penalty_cost > result.selling_price * 0.02:
        recommendations.append("⚠️ **Высокие штрафы за просрочку!** Внедрите ночную смену или ускорьте обработку заказов.")
    
    if result.recommended_model != "FBS":
        recommendations.append(f"💡 **Модель {result.recommended_model} выгоднее на {(max(result.fbo_profit, result.fbp_profit) - result.gross_profit):.0f} ₽!** Рассмотрите переход.")
    
    if result.ltv_cac_ratio < 1:
        recommendations.append("❌ **LTV/CAC < 1!** Вы тратите на привлечение клиента больше, чем он приносит. Пересмотрите маркетинговую стратегию.")
    elif result.ltv_cac_ratio < 3:
        recommendations.append("⚠️ **LTV/CAC < 3.** Работайте над удержанием клиентов и снижением стоимости привлечения.")
    
    if result.max_discount_percent < 15:
        recommendations.append("⚠️ **Малый запас для скидок!** Вы не сможете участвовать в крупных распродажах без убытка.")
    
    if not recommendations:
        recommendations.append("✅ **Отличные показатели!** Товар прибыльный, все метрики в норме. Продолжайте в том же духе!")
    
    for rec in recommendations:
        st.markdown(rec)

def show_batch_calculator():
    """Массовый расчет из файла"""
    
    st.subheader("📊 Массовый расчет FBS из файла")
    
    st.info("""
    **Загрузите файл с данными о товарах для массового расчета.**
    
    **Обязательные колонки в файле:**
    - Артикул
    - Цена продажи
    - Себестоимость
    
    **Опциональные колонки:**
    - Вес, кг
    - Длина, см
    - Ширина, см
    - Высота, см
    - Наименование
    - Категория
    """)
    
    uploaded_file = st.file_uploader(
        "📁 Загрузите файл каталога (CSV или Excel)",
        type=['csv', 'xlsx', 'xls'],
        help="Поддерживаются форматы CSV (кодировка UTF-8) и Excel (.xlsx, .xls)"
    )
    
    if uploaded_file is not None:
        try:
            # Чтение файла
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, encoding='utf-8-sig', dtype=str)
            else:
                df = pd.read_excel(uploaded_file, dtype=str)
            
            st.success(f"✅ Загружено {len(df)} товаров из файла")
            st.dataframe(df.head(10), width="stretch")
            
            # Настройка маппинга колонок
            st.markdown("### 🔧 Настройка маппинга колонок")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                artikul_col = st.selectbox(
                    "Колонка с артикулом",
                    df.columns,
                    index=next((i for i, c in enumerate(df.columns) if 'артикул' in c.lower() or 'artikul' in c.lower()), 0)
                )
            
            with col2:
                price_col = st.selectbox(
                    "Колонка с ценой продажи",
                    df.columns,
                    index=next((i for i, c in enumerate(df.columns) if 'цен' in c.lower() or 'price' in c.lower()), 0)
                )
            
            with col3:
                cost_col = st.selectbox(
                    "Колонка с себестоимостью",
                    df.columns,
                    index=next((i for i, c in enumerate(df.columns) if 'себестоимость' in c.lower() or 'cost' in c.lower() or 'закуп' in c.lower()), 0)
                )
            
            with col4:
                name_col = st.selectbox(
                    "Колонка с наименованием",
                    ["Не выбрано"] + list(df.columns),
                    index=0
                )
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                weight_col = st.selectbox("Колонка с весом (кг)", ["Не выбрано"] + list(df.columns), index=0)
            with col2:
                length_col = st.selectbox("Колонка с длиной (см)", ["Не выбрано"] + list(df.columns), index=0)
            with col3:
                width_col = st.selectbox("Колонка с шириной (см)", ["Не выбрано"] + list(df.columns), index=0)
            with col4:
                height_col = st.selectbox("Колонка с высотой (см)", ["Не выбрано"] + list(df.columns), index=0)
            
            # Общие параметры для всех товаров
            st.markdown("### ⚙️ Общие параметры для всех товаров")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                default_distance = st.number_input("Расстояние до МП (км)", value=50.0, step=1.0, key="batch_distance")
                default_transport_cost = st.number_input("Стоимость 1 км (₽)", value=20.0, step=1.0, key="batch_km_cost")
            
            with col2:
                default_pallet = st.number_input("Единиц на паллете", value=100, step=10, key="batch_pallet")
                default_packaging = st.number_input("Упаковка (₽/шт)", value=50.0, step=10.0, key="batch_packaging")
            
            with col3:
                default_pick_time = st.number_input("Время сборки (мин)", value=5.0, step=0.5, key="batch_pick_time")
                default_operator_rate = st.number_input("Ставка оператора (₽/ч)", value=300.0, step=50.0, key="batch_operator_rate")
                default_marketing = st.number_input("Маркетинг на ед. (₽)", value=100.0, step=10.0, key="batch_marketing")
            
            # Кнопка запуска расчета
            if st.button("🚀 Запустить массовый расчет", type="primary", width="stretch", key="batch_calc_button"):
                with st.spinner(f"Выполняется расчет {len(df)} товаров..."):
                    # Прогресс-бар
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    # Создание входных данных
                    input_data_list = []
                    
                    for idx, row in df.iterrows():
                        try:
                            input_data = FBSInputData(
                                artikul=str(row.get(artikul_col, f"SKU_{idx}")),
                                product_name=str(row.get(name_col, "")) if name_col != "Не выбрано" else "",
                                category="default",
                                selling_price=float(row.get(price_col, 0) or 0),
                                cogs=float(row.get(cost_col, 0) or 0),
                                weight_kg=float(row.get(weight_col, 1.0) or 1.0) if weight_col != "Не выбрано" else 1.0,
                                length_cm=float(row.get(length_col, 0) or 0) if length_col != "Не выбрано" else 0,
                                width_cm=float(row.get(width_col, 0) or 0) if width_col != "Не выбрано" else 0,
                                height_cm=float(row.get(height_col, 0) or 0) if height_col != "Не выбрано" else 0,
                                warehouse_distance_km=default_distance,
                                transport_cost_per_km=default_transport_cost,
                                pallet_capacity=default_pallet,
                                packaging_cost=default_packaging,
                                pick_pack_time_min=default_pick_time,
                                operator_hourly_rate=default_operator_rate,
                                marketing_budget_per_unit=default_marketing
                            )
                            input_data_list.append(input_data)
                        except Exception as e:
                            st.warning(f"⚠️ Ошибка в строке {idx}: {e}")
                            continue
                        
                        if idx % 100 == 0:
                            progress_bar.progress(min(idx / len(df), 1.0))
                            status_text.text(f"Подготовка данных: {idx}/{len(df)}")
                    
                    # Выполнение расчета
                    calculator = st.session_state.calculator
                    results = calculator.calculate_batch(input_data_list)
                    
                    # Сохранение результатов
                    st.session_state.results = results
                    st.session_state.input_data_list = input_data_list
                    
                    progress_bar.progress(1.0)
                    status_text.text("✅ Расчет завершен!")
                    
                    st.success(f"✅ Успешно рассчитано {len(results)} товаров!")
                    
                    # Краткая статистика
                    if results:
                        col1, col2, col3, col4 = st.columns(4)
                        
                        with col1:
                            avg_margin = np.mean([r.margin_percent for r in results])
                            st.metric("Средняя маржа", f"{avg_margin:.1f}%")
                        
                        with col2:
                            profitable = len([r for r in results if r.gross_profit > 0])
                            st.metric("Прибыльных", f"{profitable}/{len(results)}")
                        
                        with col3:
                            total_profit = sum(r.gross_profit for r in results)
                            st.metric("Общая прибыль", f"{total_profit:,.0f} ₽")
                        
                        with col4:
                            fbo_better = len([r for r in results if r.fbo_profit > r.gross_profit])
                            st.metric("FBO выгоднее", f"{fbo_better}")
        
        except Exception as e:
            st.error(f"❌ Ошибка чтения файла: {e}")
            logger.exception("Ошибка в batch calculator")

def show_tariffs_page():
    """Страница просмотра и управления тарифами"""
    
    st.markdown("## 📋 Актуальные тарифы маркетплейсов")
    
    calculator = st.session_state.calculator
    api_manager = st.session_state.api_manager
    
    # Панель управления
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        marketplace_filter = st.selectbox(
            "Маркетплейс",
            options=["Ozon", "Wildberries", "Яндекс Маркет"],
            index=["Ozon", "Wildberries", "Яндекс Маркет"].index(st.session_state.marketplace)
        )
    
    with col2:
        force_refresh = st.checkbox(
            "🔄 Принудительное обновление",
            value=False,
            help="Игнорировать кэш и загрузить свежие данные из API"
        )
    
    with col3:
        use_ai = st.checkbox(
            "🤖 Использовать DeepSeek AI",
            value=False,
            help="Если прямое API недоступно, использовать AI для получения тарифов"
        )
    
    with col4:
        if st.button("📥 Загрузить тарифы", type="primary", width="stretch"):
            with st.spinner(f"Загрузка тарифов {marketplace_filter}..."):
                calculator.set_marketplace(marketplace_filter)
                calculator.refresh_tariffs(force=force_refresh, use_ai=use_ai)
                
                st.success(f"✅ Тарифы {marketplace_filter} загружены!")
                
                if calculator.tariffs_updated_at:
                    st.info(f"🕐 Последнее обновление: {calculator.tariffs_updated_at.strftime('%d.%m.%Y %H:%M:%S')}")
                    st.info(f"📡 Источник данных: {calculator.tariffs_source}")
                
                st.rerun()
    
    # Отображение тарифов
    st.markdown("---")
    
    if calculator.current_tariffs:
        df_tariffs = api_manager.get_all_tariffs_as_dataframe(marketplace_filter)
        
        # Информация об источнике
        source_info = {}
        for cat, tariff in calculator.current_tariffs.items():
            source = tariff.get('source', 'default')
            source_info[cat] = source
        
        st.markdown(f"### Тарифы {marketplace_filter} ({len(df_tariffs)} категорий)")
        
        # Цветовое кодирование
        def color_source(val):
            if 'api' in str(val).lower():
                return 'background-color: #E0FFE0'
            elif 'deepseek' in str(val).lower():
                return 'background-color: #E0E0FF'
            else:
                return 'background-color: #FFE0E0'
        
        st.dataframe(
            df_tariffs.style.applymap(color_source, subset=['Источник']),
            width="stretch",
            height=500
        )
        
        # Статистика источников
        st.markdown("### 📊 Статистика источников данных")
        
        sources = df_tariffs['Источник'].value_counts()
        
        col1, col2, col3 = st.columns(3)
        
        with col1:
            api_count = sum(1 for v in sources.index if 'api' in str(v).lower())
            st.metric("Прямое API", api_count, delta="Наиболее точные" if api_count > 0 else "Недоступно")
        
        with col2:
            ai_count = sum(1 for v in sources.index if 'deepseek' in str(v).lower())
            st.metric("DeepSeek AI", ai_count, delta="AI-обогащение" if ai_count > 0 else "Не использовался")
        
        with col3:
            default_count = sources.get('default', 0)
            st.metric("Дефолтные значения", default_count, 
                     delta="Требуют настройки API" if default_count > 0 else "Отсутствуют",
                     delta_color="inverse" if default_count > 0 else "normal")
        
        # Тестирование API
        st.markdown("---")
        st.markdown("### 🔌 Тестирование API подключений")
        
        if st.button("🧪 Проверить все подключения", width="stretch"):
            with st.spinner("Тестирование API подключений..."):
                test_results = calculator.test_api_connections()
                
                for mp, result in test_results.items():
                    if mp == 'DeepSeek':
                        if result['status'] == 'available':
                            st.success(f"✅ {mp}: Доступен")
                        else:
                            st.warning(f"⚠️ {mp}: Не настроен")
                    else:
                        status = result.get('status', 'unknown')
                        response_time = result.get('response_time_ms', 0)
                        
                        if status == 'success':
                            st.success(f"✅ {mp}: Успешно ({response_time} мс)")
                        elif status == 'no_api_key':
                            st.warning(f"⚠️ {mp}: API ключи не настроены")
                        elif status == 'error':
                            st.error(f"❌ {mp}: Ошибка подключения — {result.get('error', 'Неизвестная ошибка')}")
                        else:
                            st.info(f"ℹ️ {mp}: {status}")
    else:
        st.warning("⚠️ Тарифы не загружены. Нажмите кнопку 'Загрузить тарифы' для получения актуальных данных.")

def show_dashboard_page():
    """Страница дашборда"""
    
    st.markdown("## 📈 Дашборд юнит-экономики")
    
    if not st.session_state.results:
        st.warning("⚠️ Нет данных для отображения. Выполните расчет в разделе 'Калькулятор FBS'.")
        return
    
    results = st.session_state.results
    visualizer = st.session_state.visualizer
    
    # Ключевые метрики
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("📦 Товаров", len(results))
    
    with col2:
        avg_margin = np.mean([r.margin_percent for r in results])
        st.metric("📊 Средняя маржа", f"{avg_margin:.1f}%")
    
    with col3:
        total_profit = sum(r.gross_profit for r in results)
        st.metric("💰 Общая прибыль", f"{total_profit:,.0f} ₽")
    
    with col4:
        profitable = len([r for r in results if r.gross_profit > 0])
        st.metric("✅ Прибыльных", f"{profitable}")
    
    with col5:
        unprofitable = len([r for r in results if r.gross_profit <= 0])
        st.metric("❌ Убыточных", f"{unprofitable}")
    
    # Графики
    st.markdown("---")
    
    if len(results) > 1:
        col1, col2 = st.columns(2)
        
        with col1:
            st.subheader("Сравнение моделей фулфилмента")
            fig_models = visualizer.create_models_comparison_chart(results)
            st.plotly_chart(fig_models, width="stretch")
        
        with col2:
            st.subheader("Точка безубыточности по расстоянию")
            fig_distance = visualizer.create_break_even_distance_chart(results)
            st.plotly_chart(fig_distance, width="stretch")
    
    # Таблица топ-10
    st.markdown("---")
    st.subheader("🏆 Топ-10 товаров по прибыли")
    
    top_results = sorted(results, key=lambda x: x.gross_profit, reverse=True)[:10]
    
    top_data = []
    for r in top_results:
        top_data.append({
            'Артикул': r.artikul,
            'Наименование': r.product_name,
            'Цена': r.selling_price,
            'Прибыль': r.gross_profit,
            'Маржа': f"{r.margin_percent:.1f}%",
            'ROI': f"{r.roi_percent:.1f}%",
            'Модель': r.recommended_model
        })
    
    df_top = pd.DataFrame(top_data)
    st.dataframe(df_top, width="stretch")

def show_export_page():
    """Страница экспорта в Excel"""
    
    st.markdown("## 📥 Экспорт результатов в Excel")
    
    if not st.session_state.results:
        st.warning("⚠️ Нет данных для экспорта. Выполните расчет в разделе 'Калькулятор FBS'.")
        return
    
    if st.session_state.exporter is None:
        st.error("❌ OpenPyXL не установлен. Выполните: pip install openpyxl")
        return
    
    results = st.session_state.results
    input_data_list = st.session_state.input_data_list
    calculator = st.session_state.calculator
    
    st.success(f"✅ Доступно для экспорта: {len(results)} товаров")
    
    st.markdown("### 📊 Профессиональный Excel-отчет")
    
    st.info("""
    **Отчет включает 8 листов:**
    1. 📋 **Тарифы МП** — актуальные тарифы из API (источник данных для формул)
    2. 📊 **Юнит-экономика FBS** — основной расчет с живыми формулами
    3. 🔄 **Сравнение моделей** — FBS vs FBO vs FBP
    4. 👥 **LTV и CAC** — метрики клиентской экономики
    5. ⚠️ **Скрытые потери FBS** — анализ неочевидных расходов
    6. 📈 **Дашборд** — сводная статистика
    7. 🔌 **API Интеграция** — информация об источниках данных
    8. 📖 **Инструкция** — руководство по использованию
    
    **💡 Главная особенность:** Формулы на листе "Юнит-экономика FBS" 
    ссылаются на лист "Тарифы МП". При обновлении тарифов в Excel 
    (измените значения в фиолетовых ячейках) — все расчеты пересчитаются автоматически!
    """)
    
    col1, col2, col3 = st.columns([1, 2, 1])
    
    with col2:
        if st.button("📥 Скачать Excel-отчет с живыми формулами", type="primary", width="stretch"):
            with st.spinner("Создание профессионального Excel-отчета..."):
                try:
                    # Создаем имя файла с датой
                    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                    filename = f"FBS_Unit_Economics_{st.session_state.marketplace}_{timestamp}.xlsx"
                    output_path = EXPORTS_DIR / filename
                    
                    # Экспорт
                    exporter = st.session_state.exporter
                    success = exporter.export_fbs_report(
                        results=results,
                        input_data_list=input_data_list,
                        calculator=calculator,
                        marketplace_name=st.session_state.marketplace,
                        output_path=str(output_path)
                    )
                    
                    if success and output_path.exists():
                        # Читаем файл для скачивания
                        with open(output_path, "rb") as f:
                            file_data = f.read()
                        
                        # Кнопка скачивания
                        st.download_button(
                            label="⬇️ Скачать Excel-отчет",
                            data=file_data,
                            file_name=filename,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_excel_button"
                        )
                        
                        st.success("""
                        ✅ **Excel-отчет успешно создан!**
                        
                        **Что дальше:**
                        1. Откройте файл в Excel
                        2. На листе "📋 Тарифы МП" проверьте актуальность данных
                        3. При необходимости измените тарифы — формулы пересчитаются автоматически
                        4. Используйте лист "🔄 Сравнение моделей" для выбора оптимальной стратегии
                        """)
                    else:
                        st.error("❌ Ошибка при создании отчета")
                
                except Exception as e:
                    st.error(f"❌ Ошибка экспорта: {e}")
                    logger.exception("Ошибка экспорта в Excel")

def show_settings_page():
    """Страница настроек"""
    
    st.markdown("## ⚙️ Настройки")
    
    tab1, tab2, tab3, tab4 = st.tabs([
        "🔑 API Ключи", 
        "🏪 Маркетплейс и налоги", 
        "🗑️ Управление данными",
        "ℹ️ О приложении"
    ])
    
    with tab1:
        st.subheader("🔑 Настройка API ключей маркетплейсов")
        
        st.info("""
        **Для автоматической загрузки актуальных тарифов необходимы API ключи:**
        
        - **Ozon**: Client ID + API Key
          - Получить: https://seller.ozon.ru/settings/api-keys
        
        - **Wildberries**: Токен продавца
          - Получить: https://suppliers.wildberries.ru/ (Настройки → API ключи)
        
        - **Яндекс Маркет**: OAuth-токен + ID кампании
          - Получить: https://partner.market.yandex.ru/ (Настройки → API)
        
        - **DeepSeek AI**: API Key
          - Получить: https://platform.deepseek.com/
        
        **Без API ключей будут использоваться дефолтные значения тарифов, которые могут быть неточными.**
        """)
        
        api_manager = st.session_state.api_manager
        
        # Ozon
        st.markdown("#### 📦 Ozon")
        col1, col2 = st.columns(2)
        
        with col1:
            ozon_client_id = st.text_input(
                "Ozon Client ID",
                value=api_manager.get_api_key('ozon_client_id') or '',
                type="password",
                key="settings_ozon_client_id"
            )
        
        with col2:
            ozon_api_key = st.text_input(
                "Ozon API Key",
                value=api_manager.get_api_key('ozon') or '',
                type="password",
                key="settings_ozon_api_key"
            )
        
        if st.button("💾 Сохранить ключи Ozon", key="save_ozon_keys"):
            api_manager.save_api_key('ozon_client_id', ozon_client_id)
            api_manager.save_api_key('ozon', ozon_api_key)
            st.success("✅ Ключи Ozon сохранены!")
        
        st.markdown("---")
        
        # Wildberries
        st.markdown("#### 📦 Wildberries")
        
        wb_token = st.text_input(
            "Wildberries API Token",
            value=api_manager.get_api_key('wildberries') or '',
            type="password",
            key="settings_wb_token"
        )
        
        if st.button("💾 Сохранить ключ Wildberries", key="save_wb_keys"):
            api_manager.save_api_key('wildberries', wb_token)
            st.success("✅ Ключ Wildberries сохранен!")
        
        st.markdown("---")
        
        # Яндекс Маркет
        st.markdown("#### 📦 Яндекс Маркет")
        
        col1, col2 = st.columns(2)
        
        with col1:
            ym_token = st.text_input(
                "Яндекс Маркет OAuth Token",
                value=api_manager.get_api_key('yandex_market') or '',
                type="password",
                key="settings_ym_token"
            )
        
        with col2:
            ym_campaign = st.text_input(
                "Campaign ID",
                value=api_manager.get_api_key('yandex_campaign_id') or '',
                key="settings_ym_campaign"
            )
        
        if st.button("💾 Сохранить ключи Яндекс Маркет", key="save_ym_keys"):
            api_manager.save_api_key('yandex_market', ym_token)
            api_manager.save_api_key('yandex_campaign_id', ym_campaign)
            st.success("✅ Ключи Яндекс Маркет сохранены!")
        
        st.markdown("---")
        
        # DeepSeek
        st.markdown("#### 🤖 DeepSeek AI")
        
        ds_key = st.text_input(
            "DeepSeek API Key",
            value=api_manager.get_api_key('deepseek') or '',
            type="password",
            key="settings_ds_key"
        )
        
        if st.button("💾 Сохранить ключ DeepSeek", key="save_ds_keys"):
            api_manager.save_api_key('deepseek', ds_key)
            st.success("✅ Ключ DeepSeek сохранен!")
    
    with tab2:
        st.subheader("🏪 Настройки маркетплейса и налогов")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### Маркетплейс по умолчанию")
            
            marketplace = st.selectbox(
                "Выберите маркетплейс",
                options=["Ozon", "Wildberries", "Яндекс Маркет"],
                index=["Ozon", "Wildberries", "Яндекс Маркет"].index(st.session_state.marketplace),
                key="settings_marketplace"
            )
            
            if st.button("💾 Сохранить маркетплейс", key="save_marketplace"):
                st.session_state.marketplace = marketplace
                st.session_state.calculator.set_marketplace(marketplace)
                st.success(f"✅ Маркетплейс '{marketplace}' сохранен!")
                st.rerun()
        
        with col2:
            st.markdown("#### Налоговая система")
            
            tax_system = st.selectbox(
                "Выберите систему налогообложения",
                options=list(TAX_SYSTEMS.keys()),
                index=list(TAX_SYSTEMS.keys()).index(st.session_state.tax_system) if st.session_state.tax_system in TAX_SYSTEMS else 0,
                key="settings_tax"
            )
            
            if st.button("💾 Сохранить налоговую систему", key="save_tax"):
                st.session_state.tax_system = tax_system
                st.session_state.calculator.tax_system = tax_system
                st.success(f"✅ Налоговая система '{tax_system}' сохранена!")
    
    with tab3:
        st.subheader("🗑️ Управление данными")
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("#### Кэш тарифов")
            
            if st.button("🗑️ Очистить кэш тарифов", width="stretch"):
                try:
                    shutil.rmtree(TARIFFS_CACHE_DIR, ignore_errors=True)
                    TARIFFS_CACHE_DIR.mkdir(exist_ok=True)
                    st.success("✅ Кэш тарифов очищен!")
                except Exception as e:
                    st.error(f"❌ Ошибка: {e}")
            
            st.caption("Кэш хранится 1 час. После очистки тарифы будут загружены заново при следующем обращении.")
        
        with col2:
            st.markdown("#### Результаты расчетов")
            
            if st.button("🗑️ Очистить все результаты", width="stretch"):
                st.session_state.results = []
                st.session_state.input_data_list = []
                st.success("✅ Результаты расчетов очищены!")
                st.rerun()
            
            st.caption(f"Сейчас хранится: {len(st.session_state.results)} результатов расчетов")
    
    with tab4:
        st.subheader("ℹ️ О приложении")
        
        st.markdown(f"""
        ### 🚀 FBS Unit Economics PRO
        
        **Версия:** {APP_VERSION}
        
        **Назначение:**
        Профессиональный калькулятор юнит-экономики для продавцов на маркетплейсах,
        работающих по модели FBS (Fulfillment by Seller).
        
        **Основные возможности:**
        - 🔌 Автоматическая загрузка тарифов через API Ozon, Wildberries, Яндекс Маркет
        - 🤖 AI-обогащение данных через DeepSeek API
        - 📊 Полный расчет юнит-экономики FBS с учетом всех скрытых расходов
        - 📏 Расчет точки безубыточности по расстоянию First Mile
        - 👥 Расчет LTV, CAC и других клиентских метрик
        - 🔄 Сравнение моделей FBS, FBO, FBP
        - 📥 Экспорт в Excel с живыми формулами и листом тарифов
        
        **Технологии:**
        - Python 3.12
        - Streamlit
        - Pandas, NumPy
        - Plotly (визуализация)
        - OpenPyXL (Excel экспорт)
        - Cryptography (шифрование ключей)
        
        **API Интеграции:**
        - Ozon Seller API
        - Wildberries Supplier API
        - Яндекс Маркет Partner API
        - DeepSeek AI API
        """)

# ============================================================================
# БЛОК 11: ГЛАВНАЯ ФУНКЦИЯ ПРИЛОЖЕНИЯ
# ============================================================================

def main():
    """Главная функция запуска приложения"""
    
    # Настройка страницы Streamlit
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="🚀",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Инициализация состояний
    init_session_state()
    
    # Отрисовка боковой панели
    render_sidebar()
    
    # Маршрутизация по разделам
    current_section = st.session_state.get('current_section', 'main')
    
    if current_section == 'main':
        show_main_page()
    elif current_section == 'calculator':
        show_calculator_page()
    elif current_section == 'tariffs':
        show_tariffs_page()
    elif current_section == 'dashboard':
        show_dashboard_page()
    elif current_section == 'export':
        show_export_page()
    elif current_section == 'settings':
        show_settings_page()
    else:
        show_main_page()

# Точка входа
if __name__ == "__main__":
    main()
