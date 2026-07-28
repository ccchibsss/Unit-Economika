"""
============================================================================
🚀 FBS UNIT ECONOMICS PRO 2026 — МОНОЛИТНОЕ ПРИЛОЖЕНИЕ
============================================================================
Профессиональный калькулятор юнит-экономики для FBS-модели
Маркетплейсы: Ozon, Wildberries, Яндекс Маркет
Автор: Операционный директор с экспертизой в FBS
Версия: 4.0.0

ОСНОВНЫЕ ВОЗМОЖНОСТИ:
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
from functools import lru_cache, wraps
import uuid
import math
import warnings
import io
from concurrent.futures import ThreadPoolExecutor, as_completed
import traceback

# Попытка импорта дополнительных библиотек
try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False
    Fernet = None

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, NamedStyle, numbers
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule, FormulaRule
    from openpyxl.chart import BarChart, Reference, PieChart, LineChart
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.series import DataPoint
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.drawing.image import Image
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    import gspread
    from google.oauth2.service_account import Credentials
    from google.oauth2 import service_account
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False
    gspread = None

warnings.filterwarnings('ignore')

# ============================================================================
# БЛОК 0: БАЗОВАЯ КОНФИГУРАЦИЯ И НАСТРОЙКИ
# ============================================================================

# Версия приложения
APP_VERSION = "4.0.0"
APP_NAME = "🚀 FBS Юнит-экономика PRO 2026"
APP_DESCRIPTION = "Профессиональный расчет юнит-экономики для FBS-модели на Ozon, Wildberries, Яндекс Маркет"

# Настройка путей
BASE_DIR = Path(__file__).parent.resolve() if '__file__' in dir() else Path.cwd()
DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
LOGS_DIR = BASE_DIR / "logs"
EXPORTS_DIR = BASE_DIR / "exports"
CONFIG_DIR = BASE_DIR / "config"
TEMP_DIR = BASE_DIR / "temp"

# Создание директорий
for dir_path in [DATA_DIR, CACHE_DIR, LOGS_DIR, EXPORTS_DIR, CONFIG_DIR, TEMP_DIR]:
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
# БЛОК 1: КОНФИГУРАЦИИ МАРКЕТПЛЕЙСОВ (АКТУАЛЬНЫЕ ТАРИФЫ 2026)
# ============================================================================

@dataclass
class MarketplaceConfig:
    """Конфигурация маркетплейса со всеми тарифами"""
    name: str
    commission_rates: Dict[str, float]  # Комиссии по категориям
    min_commission: float  # Минимальная комиссия
    last_mile_base: float  # Базовая стоимость последней мили
    last_mile_per_kg: float  # Доплата за кг последней мили
    last_mile_per_km: float  # Доплата за км последней мили
    acquiring_fee: float  # Эквайринг
    return_fee: float  # Комиссия за возврат
    penalty_rate: float  # Штраф за просрочку (% от цены)
    penalty_time_hours: int  # Время на передачу заказа (часы)
    fbo_multiplier: float  # Множитель для FBO (экономия на логистике)
    fbp_multiplier: float  # Множитель для FBP
    storage_base_rate: float  # Базовая ставка хранения (для FBO/FBP анализа)
    min_logistics: float  # Минимальная стоимость логистики

# Конфигурация Ozon FBS
OZON_FBS_CONFIG = MarketplaceConfig(
    name="Ozon",
    commission_rates={
        "default": 0.15,
        "auto_parts": 0.12,
        "electronics": 0.10,
        "clothing": 0.20,
        "home": 0.15,
        "sport": 0.15,
        "beauty": 0.18,
        "books": 0.15,
        "toys": 0.15,
        "food": 0.10
    },
    min_commission=30.0,
    last_mile_base=50.0,
    last_mile_per_kg=15.0,
    last_mile_per_km=3.5,
    acquiring_fee=0.015,
    return_fee=0.02,
    penalty_rate=0.05,
    penalty_time_hours=24,
    fbo_multiplier=0.75,
    fbp_multiplier=0.60,
    storage_base_rate=0.30,
    min_logistics=25.0
)

# Конфигурация Wildberries FBS
WILDBERRIES_FBS_CONFIG = MarketplaceConfig(
    name="Wildberries",
    commission_rates={
        "default": 0.16,
        "auto_parts": 0.13,
        "electronics": 0.11,
        "clothing": 0.21,
        "home": 0.16,
        "sport": 0.16,
        "beauty": 0.19,
        "books": 0.16,
        "toys": 0.16,
        "food": 0.11
    },
    min_commission=28.0,
    last_mile_base=45.0,
    last_mile_per_kg=14.0,
    last_mile_per_km=3.2,
    acquiring_fee=0.015,
    return_fee=0.018,
    penalty_rate=0.08,
    penalty_time_hours=24,
    fbo_multiplier=0.70,
    fbp_multiplier=0.55,
    storage_base_rate=0.25,
    min_logistics=22.0
)

# Конфигурация Яндекс Маркет FBS
YANDEX_FBS_CONFIG = MarketplaceConfig(
    name="Яндекс Маркет",
    commission_rates={
        "default": 0.145,
        "auto_parts": 0.11,
        "electronics": 0.09,
        "clothing": 0.19,
        "home": 0.145,
        "sport": 0.145,
        "beauty": 0.175,
        "books": 0.145,
        "toys": 0.145,
        "food": 0.09
    },
    min_commission=35.0,
    last_mile_base=55.0,
    last_mile_per_kg=16.0,
    last_mile_per_km=3.8,
    acquiring_fee=0.015,
    return_fee=0.025,
    penalty_rate=0.07,
    penalty_time_hours=24,
    fbo_multiplier=0.80,
    fbp_multiplier=0.65,
    storage_base_rate=0.35,
    min_logistics=30.0
)

# Словарь конфигураций
MARKETPLACE_CONFIGS = {
    "Ozon": OZON_FBS_CONFIG,
    "Wildberries": WILDBERRIES_FBS_CONFIG,
    "Яндекс Маркет": YANDEX_FBS_CONFIG
}

# Налоговые системы
TAX_SYSTEMS = {
    "УСН 6% (доходы)": {"rate": 0.06, "base": "revenue", "name": "УСН_6"},
    "УСН 15% (доходы-расходы)": {"rate": 0.15, "base": "profit", "min_rate": 0.01, "name": "УСН_15"},
    "ОСН (общая)": {"rate": 0.20, "base": "profit", "name": "ОСН"},
    "НПД (самозанятый)": {"rate": 0.06, "base": "revenue", "name": "НПД"},
    "Патент": {"rate": 0.06, "base": "revenue", "name": "Патент"}
}

# ============================================================================
# БЛОК 2: ДЕКОРАТОРЫ И УТИЛИТЫ
# ============================================================================

def timing_decorator(func):
    """Декоратор для измерения времени выполнения"""
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
    """Декоратор для повторных попыток при ошибках"""
    def decorator(func):
        @wraps(func)
        def wrapper(*args, **kwargs):
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except Exception as e:
                    if attempt == max_retries - 1:
                        raise
                    logger.warning(f"⚠️ Попытка {attempt + 1} для {func.__name__} не удалась: {e}")
                    time.sleep(delay)
            return None
        return wrapper
    return decorator

def memoize(func):
    """Мемоизация для кэширования результатов"""
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
        self.progress = 0
        self.status = ""
        self.total = 0
        self.current = 0
    
    def update(self, current: int, total: int, status: str = ""):
        self.current = current
        self.total = total
        self.progress = min(current / total, 1.0) if total > 0 else 0
        self.status = status
    
    def get_progress(self) -> float:
        return self.progress
    
    def get_status(self) -> str:
        return self.status

# ============================================================================
# БЛОК 3: БЕЗОПАСНОЕ ХРАНЕНИЕ ДАННЫХ (ШИФРОВАНИЕ)
# ============================================================================

class SecureDataManager:
    """
    Менеджер безопасного хранения конфиденциальных данных.
    Использует Fernet для шифрования.
    """
    def __init__(self):
        self.key_file = CONFIG_DIR / ".master_key"
        self.data_file = CONFIG_DIR / ".secure_data.enc"
        self._fernet = None
        self._init_encryption()
    
    def _init_encryption(self):
        """Инициализация шифрования"""
        if not CRYPTO_AVAILABLE:
            logger.warning("⚠️ Cryptography не установлен. Данные не будут зашифрованы.")
            return
        
        try:
            if self.key_file.exists():
                key = self.key_file.read_bytes()
            else:
                key = Fernet.generate_key()
                self.key_file.write_bytes(key)
                # Устанавливаем права только для владельца
                os.chmod(self.key_file, 0o600)
            
            self._fernet = Fernet(key)
            logger.info("✅ Шифрование инициализировано")
        except Exception as e:
            logger.error(f"❌ Ошибка инициализации шифрования: {e}")
            self._fernet = None
    
    def save_data(self, data: Dict[str, Any]) -> bool:
        """Сохраняет зашифрованные данные"""
        if not self._fernet:
            return False
        
        try:
            json_data = json.dumps(data, ensure_ascii=False)
            encrypted = self._fernet.encrypt(json_data.encode('utf-8'))
            self.data_file.write_bytes(encrypted)
            os.chmod(self.data_file, 0o600)
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка сохранения данных: {e}")
            return False
    
    def load_data(self) -> Dict[str, Any]:
        """Загружает и расшифровывает данные"""
        if not self._fernet or not self.data_file.exists():
            return {}
        
        try:
            encrypted = self.data_file.read_bytes()
            decrypted = self._fernet.decrypt(encrypted)
            return json.loads(decrypted.decode('utf-8'))
        except Exception as e:
            logger.error(f"❌ Ошибка загрузки данных: {e}")
            return {}
    
    def delete_data(self) -> bool:
        """Удаляет зашифрованные данные"""
        try:
            if self.data_file.exists():
                self.data_file.unlink()
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка удаления данных: {e}")
            return False

# ============================================================================
# БЛОК 4: КЭШИРОВАНИЕ И УПРАВЛЕНИЕ ПАМЯТЬЮ
# ============================================================================

class CacheManager:
    """Менеджер кэширования для оптимизации производительности"""
    def __init__(self, max_size_mb: int = 500):
        self.cache_dir = CACHE_DIR
        self.max_size_mb = max_size_mb
        self._memory_cache: Dict[str, Any] = {}
        self._cache_timestamps: Dict[str, float] = {}
        self._cache_ttl = 3600  # 1 час
    
    @timing_decorator
    def get_or_compute(self, key: str, compute_func: Callable, *args, **kwargs) -> Any:
        """Получает значение из кэша или вычисляет"""
        # Проверка memory cache
        if key in self._memory_cache:
            if time.time() - self._cache_timestamps.get(key, 0) < self._cache_ttl:
                logger.debug(f"📦 Кэш попадание: {key}")
                return self._memory_cache[key]
        
        # Проверка disk cache
        disk_result = self._load_from_disk(key)
        if disk_result is not None:
            self._memory_cache[key] = disk_result
            self._cache_timestamps[key] = time.time()
            return disk_result
        
        # Вычисление
        logger.debug(f"🔄 Вычисление: {key}")
        result = compute_func(*args, **kwargs)
        
        # Сохранение в кэш
        self._memory_cache[key] = result
        self._cache_timestamps[key] = time.time()
        self._save_to_disk(key, result)
        
        return result
    
    def _load_from_disk(self, key: str) -> Optional[Any]:
        """Загрузка из дискового кэша"""
        cache_file = self.cache_dir / f"{hashlib.md5(key.encode()).hexdigest()}.cache"
        if cache_file.exists():
            try:
                with open(cache_file, 'rb') as f:
                    data = pickle.load(f)
                return data
            except Exception:
                pass
        return None
    
    def _save_to_disk(self, key: str, data: Any):
        """Сохранение в дисковый кэш"""
        try:
            cache_file = self.cache_dir / f"{hashlib.md5(key.encode()).hexdigest()}.cache"
            with open(cache_file, 'wb') as f:
                pickle.dump(data, f)
        except Exception as e:
            logger.debug(f"Не удалось сохранить кэш на диск: {e}")
    
    def clear_cache(self):
        """Очистка кэша"""
        self._memory_cache.clear()
        self._cache_timestamps.clear()
        for cache_file in self.cache_dir.glob("*.cache"):
            cache_file.unlink()
        logger.info("🗑️ Кэш очищен")

# ============================================================================
# БЛОК 5: ОСНОВНОЙ КАЛЬКУЛЯТОР FBS ЮНИТ-ЭКОНОМИКИ
# ============================================================================

@dataclass
class FBSInputData:
    """Входные данные для расчета FBS"""
    # Основные параметры
    artikul: str = ""
    product_name: str = ""
    category: str = "default"
    
    # Финансы
    selling_price: float = 0.0  # Цена продажи
    cogs: float = 0.0  # Себестоимость закупки
    
    # Физические параметры
    weight_kg: float = 0.0  # Вес брутто
    length_cm: float = 0.0  # Длина
    width_cm: float = 0.0  # Ширина
    height_cm: float = 0.0  # Высота
    
    # FBS специфика
    first_mile_cost_per_unit: float = 0.0  # Стоимость доставки до склада МП на 1 шт
    packaging_cost: float = 0.0  # Стоимость упаковочного материала
    pick_pack_time_min: float = 5.0  # Время сборки заказа (минуты)
    operator_hourly_rate: float = 300.0  # Ставка оператора сборки (₽/час)
    warehouse_distance_km: float = 0.0  # Расстояние до склада МП
    
    # Логистика
    transport_type: str = "own"  # Тип транспорта: own, cdek, delovye_linii
    transport_cost_per_km: float = 20.0  # Стоимость 1 км транспорта
    pallet_capacity: int = 100  # Количество единиц на паллете
    pallet_cost: float = 2000.0  # Стоимость паллета
    
    # Маркетинг
    marketing_budget_per_unit: float = 0.0  # Рекламный бюджет на единицу
    
    # Складские параметры
    stock_depth_days: int = 30  # Глубина запаса
    daily_sales: int = 5  # Продаж в день
    warehouse_rent_per_sqm: float = 500.0  # Аренда склада за м²
    warehouse_space_per_unit: float = 0.01  # Место на складе на единицу (м²)
    
    # Повторные продажи
    repeat_purchase_rate: float = 0.3  # Коэффициент повторных покупок
    avg_purchases_per_year: float = 2.5  # Среднее количество покупок в год
    customer_retention_rate: float = 0.7  # Коэффициент удержания (CRR)
    discount_rate: float = 0.1  # Ставка дисконтирования
    
    # Режим работы
    has_night_shift: bool = False  # Наличие ночной смены
    processing_capacity_per_hour: int = 20  # Пропускная способность обработки заказов в час
    
    def to_dict(self) -> Dict[str, Any]:
        """Преобразование в словарь"""
        return asdict(self)

@dataclass
class FBSResultData:
    """Результаты расчета FBS юнит-экономики"""
    # Основные метрики
    artikul: str = ""
    product_name: str = ""
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
    penalty_probability: float = 0.0  # Вероятность просрочки
    break_even_distance_km: float = 0.0  # Точка безубыточности по расстоянию
    max_discount_percent: float = 0.0  # Максимальная скидка
    safety_margin_price: float = 0.0  # Запас прочности по цене
    
    # LTV и CAC
    ltv: float = 0.0
    cac: float = 0.0
    ltv_cac_ratio: float = 0.0
    
    # Сравнение с другими моделями
    fbo_profit: float = 0.0
    fbp_profit: float = 0.0
    recommended_model: str = ""
    
    def to_dict(self) -> Dict[str, Any]:
        """Преобразование в словарь"""
        return asdict(self)

class FBSUnitEconomicsCalculator:
    """
    Профессиональный калькулятор юнит-экономики для FBS-модели.
    Учитывает специфику FBS: двойную логистику, штрафы, Pick & Pack.
    """
    
    def __init__(self, marketplace_config: MarketplaceConfig = None, tax_system: str = "УСН 6% (доходы)"):
        self.marketplace_config = marketplace_config or OZON_FBS_CONFIG
        self.tax_system = tax_system
        self.cache_manager = CacheManager()
        self.progress_tracker = ProgressTracker()
        
        # Настройки по умолчанию
        self.default_pick_pack_time = 5.0  # минут
        self.default_operator_rate = 300.0  # ₽/час
        self.default_first_mile_per_km = 20.0  # ₽/км
        self.default_penalty_probability_no_night = 0.35  # Без ночной смены
        self.default_penalty_probability_with_night = 0.05  # С ночной сменой
    
    def set_marketplace(self, marketplace_name: str):
        """Установка маркетплейса"""
        if marketplace_name in MARKETPLACE_CONFIGS:
            self.marketplace_config = MARKETPLACE_CONFIGS[marketplace_name]
            logger.info(f"🏪 Установлен маркетплейс: {marketplace_name}")
        else:
            logger.warning(f"⚠️ Неизвестный маркетплейс: {marketplace_name}")
    
    @timing_decorator
    def calculate_unit_economics(self, input_data: FBSInputData) -> FBSResultData:
        """
        Основной расчет юнит-экономики FBS.
        Возвращает полную детализацию всех расходов и метрик.
        """
        result = FBSResultData()
        result.artikul = input_data.artikul
        result.product_name = input_data.product_name
        result.selling_price = input_data.selling_price
        
        config = self.marketplace_config
        
        # === 1. КОМИССИЯ МАРКЕТПЛЕЙСА ===
        commission_rate = config.commission_rates.get(input_data.category, config.commission_rates["default"])
        result.commission = max(
            input_data.selling_price * commission_rate,
            config.min_commission
        )
        
        # === 2. FIRST MILE (ВАША ЛОГИСТИКА ДО СКЛАДА МП) ===
        if input_data.first_mile_cost_per_unit > 0:
            result.first_mile_cost = input_data.first_mile_cost_per_unit
        else:
            # Расчет на основе расстояния и типа транспорта
            pallet_units = max(input_data.pallet_capacity, 1)
            cost_per_pallet = input_data.warehouse_distance_km * input_data.transport_cost_per_km
            result.first_mile_cost = cost_per_pallet / pallet_units
        
        # === 3. LAST MILE (ЛОГИСТИКА МАРКЕТПЛЕЙСА) ===
        # Расчет оплачиваемого веса
        vol_weight = (input_data.length_cm * input_data.width_cm * input_data.height_cm) / 5000.0
        billable_weight = max(input_data.weight_kg, vol_weight)
        billable_weight = math.ceil(billable_weight * 2) / 2  # Округление до 0.5
        
        result.last_mile_cost = max(
            config.last_mile_base + (billable_weight * config.last_mile_per_kg),
            config.min_logistics
        )
        
        # === 4. PICK & PACK (СТОИМОСТЬ ОБРАБОТКИ ЗАКАЗА) ===
        pick_pack_hours = input_data.pick_pack_time_min / 60.0
        result.pick_pack_cost = pick_pack_hours * input_data.operator_hourly_rate
        
        # === 5. УПАКОВКА ===
        result.packaging_cost = input_data.packaging_cost
        
        # === 6. ЭКВАЙРИНГ ===
        result.acquiring_cost = input_data.selling_price * config.acquiring_fee
        
        # === 7. ВОЗВРАТЫ ===
        result.return_cost = input_data.selling_price * config.return_fee
        
        # === 8. ШТРАФЫ ЗА ПРОСРОЧКУ (PENALTY RATE) ===
        # Расчет вероятности просрочки
        if input_data.has_night_shift:
            penalty_prob = self.default_penalty_probability_with_night
        else:
            # Без ночной смены: заказы после 18:00 имеют высокий риск просрочки
            penalty_prob = self.default_penalty_probability_no_night
        
        result.penalty_probability = penalty_prob
        result.penalty_cost = input_data.selling_price * config.penalty_rate * penalty_prob
        
        # === 9. МАРКЕТИНГ ===
        result.marketing_cost = input_data.marketing_budget_per_unit
        
        # === 10. СКЛАДСКИЕ РАСХОДЫ (РАСПРЕДЕЛЕНИЕ НА ЕДИНИЦУ) ===
        total_stock = input_data.stock_depth_days * input_data.daily_sales
        monthly_rent = input_data.warehouse_rent_per_sqm * input_data.warehouse_space_per_unit * total_stock
        result.warehouse_cost = monthly_rent / (30 * input_data.daily_sales) if input_data.daily_sales > 0 else 0
        
        # === 11. НАЛОГ ===
        tax_config = TAX_SYSTEMS.get(self.tax_system, TAX_SYSTEMS["УСН 6% (доходы)"])
        if tax_config["base"] == "revenue":
            result.tax_cost = input_data.selling_price * tax_config["rate"]
        else:
            # Налог с прибыли (предварительный расчет)
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
        
        # === 12. ИТОГО РАСХОДОВ И ПРИБЫЛЬ ===
        result.total_expenses = (
            input_data.cogs + result.commission + result.first_mile_cost +
            result.last_mile_cost + result.pick_pack_cost + result.packaging_cost +
            result.acquiring_cost + result.return_cost + result.penalty_cost +
            result.marketing_cost + result.warehouse_cost + result.tax_cost
        )
        
        result.gross_profit = result.selling_price - result.total_expenses
        result.margin_percent = (result.gross_profit / result.selling_price * 100) if result.selling_price > 0 else 0
        result.roi_percent = (result.gross_profit / input_data.cogs * 100) if input_data.cogs > 0 else 0
        
        # === 13. ТОЧКА БЕЗУБЫТОЧНОСТИ ПО РАССТОЯНИЮ (FIRST MILE) ===
        # Максимальное расстояние, при котором First Mile окупается
        if result.first_mile_cost > 0:
            cost_per_km = input_data.transport_cost_per_km / max(input_data.pallet_capacity, 1)
            result.break_even_distance_km = result.gross_profit / cost_per_km if cost_per_km > 0 else float('inf')
        else:
            result.break_even_distance_km = float('inf')
        
        # === 14. ЗАПАС ПРОЧНОСТИ ПО ЦЕНЕ (ДЛЯ РАСПРОДАЖ) ===
        # На сколько можно снизить цену, чтобы остаться в плюсе
        variable_costs_percent = (
            commission_rate + config.acquiring_fee + config.return_fee +
            config.penalty_rate * penalty_prob
        )
        fixed_costs = (
            input_data.cogs + result.first_mile_cost + result.last_mile_cost +
            result.pick_pack_cost + result.packaging_cost + result.marketing_cost +
            result.warehouse_cost
        )
        
        # Расчет минимальной цены (точка безубыточности)
        if (1 - variable_costs_percent - TAX_SYSTEMS[self.tax_system]["rate"]) > 0:
            min_price = fixed_costs / (1 - variable_costs_percent - TAX_SYSTEMS[self.tax_system]["rate"])
        else:
            min_price = fixed_costs * 1.5
        
        result.safety_margin_price = input_data.selling_price - min_price
        result.max_discount_percent = ((input_data.selling_price - min_price) / input_data.selling_price * 100) if input_data.selling_price > 0 else 0
        
        # === 15. LTV И CAC ===
        # LTV = (Средний чек * Кол-во повторных покупок * CRR) / (1 + Дисконт)
        result.ltv = (
            input_data.selling_price * 
            input_data.avg_purchases_per_year * 
            input_data.customer_retention_rate
        ) / (1 + input_data.discount_rate)
        
        # CAC = (Маркетинг + Штрафы + First Mile) / Новые клиенты
        total_acquisition_cost = (
            result.marketing_cost + result.penalty_cost + result.first_mile_cost
        )
        # Предполагаем, что 30% покупателей - новые
        new_customers_per_order = 0.3
        result.cac = total_acquisition_cost / new_customers_per_order if new_customers_per_order > 0 else 0
        
        result.ltv_cac_ratio = result.ltv / result.cac if result.cac > 0 else float('inf')
        
        # === 16. СРАВНЕНИЕ С FBO И FBP ===
        # FBO: экономим на First Mile, но платим за хранение
        fbo_storage_days = 30  # Среднее хранение при FBO
        storage_volume = (input_data.length_cm * input_data.width_cm * input_data.height_cm) / 1000000  # м³
        fbo_storage_cost = storage_volume * config.storage_base_rate * 30 * fbo_storage_days
        
        fbo_commission = result.commission  # Та же комиссия
        fbo_logistics = result.last_mile_cost * config.fbo_multiplier  # Логистика дешевле
        
        fbo_expenses = (
            input_data.cogs + fbo_commission + fbo_logistics + fbo_storage_cost +
            result.acquiring_cost + result.return_cost + result.tax_cost +
            result.marketing_cost + result.packaging_cost
        )
        result.fbo_profit = input_data.selling_price - fbo_expenses
        
        # FBP: частичный фулфилмент
        fbp_logistics = result.last_mile_cost * config.fbp_multiplier
        fbp_expenses = (
            input_data.cogs + fbo_commission + fbp_logistics + fbo_storage_cost * 0.5 +
            result.acquiring_cost + result.return_cost + result.tax_cost +
            result.marketing_cost + result.packaging_cost
        )
        result.fbp_profit = input_data.selling_price - fbp_expenses
        
        # Рекомендация модели
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
        Поддерживает параллельную обработку для больших объемов.
        """
        results = []
        total = len(input_data_list)
        
        if total > 100 and use_parallel:
            # Параллельная обработка для больших партий
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                futures = {executor.submit(self.calculate_unit_economics, data): i 
                          for i, data in enumerate(input_data_list)}
                
                for future in as_completed(futures):
                    try:
                        result = future.result()
                        results.append((futures[future], result))
                    except Exception as e:
                        logger.error(f"❌ Ошибка расчета: {e}")
                
                # Сортировка по исходному порядку
                results.sort(key=lambda x: x[0])
                results = [r[1] for r in results]
        else:
            # Последовательная обработка для небольших партий
            for i, data in enumerate(input_data_list):
                try:
                    result = self.calculate_unit_economics(data)
                    results.append(result)
                except Exception as e:
                    logger.error(f"❌ Ошибка расчета для {data.artikul}: {e}")
                
                if i % 100 == 0:
                    self.progress_tracker.update(i, total, f"Обработано {i}/{total}")
        
        self.progress_tracker.update(total, total, "Расчет завершен")
        return results

# ============================================================================
# БЛОК 6: ЭКСПОРТ В EXCEL С ЖИВЫМИ ФОРМУЛАМИ
# ============================================================================

class ProfessionalExcelExporter:
    """
    Профессиональный экспорт в Excel с живыми формулами, 
    визуализацией и интерактивными элементами.
    Поддерживает большие объемы данных (до 1 000 000 строк).
    """
    
    def __init__(self):
        if not OPENPYXL_AVAILABLE:
            raise ImportError("openpyxl не установлен")
        
        # Стили
        self.header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
        self.header_font = Font(bold=True, color="FFFFFF", size=11, name="Arial")
        self.header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        
        self.input_fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
        self.formula_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        self.result_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        self.profit_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        self.loss_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        
        self.thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        self.title_font = Font(bold=True, size=14, name="Arial", color="1a1a2e")
        self.subtitle_font = Font(bold=True, size=12, name="Arial", color="333333")
    
    @timing_decorator
    def export_fbs_report(self, results: List[FBSResultData], 
                         input_data_list: List[FBSInputData],
                         marketplace_name: str,
                         output_path: str) -> bool:
        """
        Создание профессионального Excel-отчета с живыми формулами.
        """
        try:
            wb = Workbook()
            
            # Удаляем стандартный лист
            if 'Sheet' in wb.sheetnames:
                del wb['Sheet']
            
            # === ЛИСТ 1: ЮНИТ-ЭКОНОМИКА FBS (ОСНОВНОЙ) ===
            ws_main = wb.create_sheet("📊 Юнит-экономика FBS", 0)
            self._create_main_sheet(ws_main, results, input_data_list, marketplace_name)
            
            # === ЛИСТ 2: СКРЫТЫЕ ПОТЕРИ FBS ===
            ws_hidden = wb.create_sheet("⚠️ Скрытые потери FBS")
            self._create_hidden_losses_sheet(ws_hidden, results, input_data_list, marketplace_name)
            
            # === ЛИСТ 3: АНАЛИЗ МОДЕЛЕЙ (FBS vs FBO vs FBP) ===
            ws_models = wb.create_sheet("🔄 Сравнение моделей")
            self._create_models_comparison_sheet(ws_models, results, input_data_list, marketplace_name)
            
            # === ЛИСТ 4: LTV И CAC ===
            ws_ltv = wb.create_sheet("👥 LTV и CAC")
            self._create_ltv_cac_sheet(ws_ltv, results, input_data_list, marketplace_name)
            
            # === ЛИСТ 5: ДАШБОРД ===
            ws_dashboard = wb.create_sheet("📈 Дашборд")
            self._create_dashboard_sheet(ws_dashboard, results, marketplace_name)
            
            # === ЛИСТ 6: РЕКОМЕНДАЦИИ ===
            ws_recommendations = wb.create_sheet("💡 Рекомендации")
            self._create_recommendations_sheet(ws_recommendations, results, marketplace_name)
            
            # === ЛИСТ 7: ИНСТРУКЦИЯ ===
            ws_instructions = wb.create_sheet("📖 Инструкция")
            self._create_instructions_sheet(ws_instructions)
            
            # Сохранение
            wb.save(output_path)
            logger.info(f"✅ Excel отчет сохранен: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"❌ Ошибка создания Excel отчета: {e}")
            traceback.print_exc()
            return False
    
    def _create_main_sheet(self, ws, results: List[FBSResultData], 
                          input_data_list: List[FBSInputData], marketplace_name: str):
        """Создание основного листа с юнит-экономикой"""
        
        # Заголовок
        ws.merge_cells('A1:AH1')
        title_cell = ws.cell(row=1, column=1, 
                            value=f"🚀 Юнит-экономика FBS - {marketplace_name} - {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        title_cell.font = self.title_font
        title_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws.row_dimensions[1].height = 40
        
        # Подзаголовок с легендой
        ws.merge_cells('A2:AH2')
        legend_cell = ws.cell(row=2, column=1,
                             value="🟡 Вводные данные | 🟢 Формулы/Расчеты | 🔵 Итоговые метрики")
        legend_cell.font = Font(size=10, italic=True, color="666666")
        legend_cell.alignment = Alignment(horizontal="center")
        
        # === ЗАГОЛОВКИ КОЛОНОК ===
        headers = [
            # Вводные данные (A-H)
            ("Артикул", 15), ("Наименование", 25), ("Категория", 15),
            ("Цена продажи, ₽", 15), ("Себестоимость, ₽", 15),
            ("Вес, кг", 10), ("Длина, см", 10), ("Ширина, см", 10),
            # Вводные данные продолжение (I-P)
            ("Высота, см", 10), ("Расстояние до МП, км", 15),
            ("Стоимость 1 км, ₽", 15), ("Единиц на паллете", 15),
            ("Упаковка, ₽", 12), ("Время сборки, мин", 12),
            ("Ставка оператора, ₽/ч", 15), ("Маркетинг, ₽", 12),
            # Расчетные параметры (Q-W)
            ("Комиссия, %", 12), ("Штраф за просрочку, %", 15),
            ("Эквайринг, %", 12), ("Возвраты, %", 12),
            ("Налог, %", 10), ("Объемный вес, кг", 12),
            ("Оплачиваемый вес, кг", 15),
            # Расходы (X-AF)
            ("Комиссия, ₽", 12), ("First Mile, ₽", 12),
            ("Last Mile, ₽", 12), ("Pick & Pack, ₽", 12),
            ("Упаковка (расчет), ₽", 14), ("Эквайринг, ₽", 12),
            ("Возвраты, ₽", 12), ("Штрафы, ₽", 12),
            ("Маркетинг (расчет), ₽", 15), ("Склад, ₽", 12),
            ("Налог, ₽", 12),
            # Итоги (AG-AH)
            ("ИТОГО расходов, ₽", 15), ("ПРИБЫЛЬ, ₽", 15)
        ]
        
        # Дополнительные колонки для формул
        extra_headers = [
            ("МАРЖА, %", 12), ("ROI, %", 12),
            ("Мин. цена, ₽", 12), ("Макс. скидка, %", 12),
            ("Точка беззуб. км", 15), ("LTV, ₽", 12),
            ("CAC, ₽", 12), ("LTV/CAC", 10),
            ("Прибыль FBO, ₽", 15), ("Прибыль FBP, ₽", 15),
            ("Рек. модель", 12)
        ]
        
        all_headers = headers + extra_headers
        
        # Запись заголовков
        for col_idx, (header_text, width) in enumerate(all_headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header_text)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.thin_border
            ws.column_dimensions[get_column_letter(col_idx)].width = width
        
        ws.row_dimensions[3].height = 35
        
        # Заморозка панелей
        ws.freeze_panes = "A4"
        
        # === ДАННЫЕ И ФОРМУЛЫ ===
        for row_idx, (result, input_data) in enumerate(zip(results, input_data_list), 4):
            # --- ВВОДНЫЕ ДАННЫЕ (желтые) ---
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
            
            # --- ПАРАМЕТРЫ СТАВОК (желтые) ---
            param_fields = [
                (17, self.marketplace_config.commission_rates.get(input_data.category, 
                                   self.marketplace_config.commission_rates["default"]) * 100),
                (18, self.marketplace_config.penalty_rate * 100),
                (19, self.marketplace_config.acquiring_fee * 100),
                (20, self.marketplace_config.return_fee * 100),
                (21, TAX_SYSTEMS[self.tax_system]["rate"] * 100)
            ]
            
            for col, value in param_fields:
                cell = ws.cell(row=row_idx, column=col, value=value)
                cell.fill = self.input_fill
                cell.border = self.thin_border
                cell.number_format = '0.00"%"'
            
            # --- РАСЧЕТНЫЕ ФОРМУЛЫ (зеленые) ---
            # Объемный вес (колонка 22 = V)
            formula_vol_weight = f"=IF(F{row_idx}*G{row_idx}*H{row_idx}>0, (G{row_idx}*H{row_idx}*I{row_idx})/5000, 0)"
            cell = ws.cell(row=row_idx, column=22, value=formula_vol_weight)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Оплачиваемый вес (колонка 23 = W)
            formula_billable = f"=CEILING(MAX(F{row_idx}, V{row_idx}), 0.5)"
            cell = ws.cell(row=row_idx, column=23, value=formula_billable)
            cell.fill = self.formula_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # --- РАСХОДЫ (синие) ---
            # Комиссия (колонка 24 = X)
            formula_commission = f"=MAX(D{row_idx}*Q{row_idx}/100, {self.marketplace_config.min_commission})"
            cell = ws.cell(row=row_idx, column=24, value=formula_commission)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # First Mile (колонка 25 = Y)
            formula_first_mile = f"=IF(L{row_idx}>0, (J{row_idx}*K{row_idx})/L{row_idx}, M{row_idx})"
            cell = ws.cell(row=row_idx, column=25, value=formula_first_mile)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Last Mile (колонка 26 = Z)
            formula_last_mile = f"=MAX({self.marketplace_config.last_mile_base}+W{row_idx}*{self.marketplace_config.last_mile_per_kg}, {self.marketplace_config.min_logistics})"
            cell = ws.cell(row=row_idx, column=26, value=formula_last_mile)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Pick & Pack (колонка 27 = AA)
            formula_pick_pack = f"=(N{row_idx}/60)*O{row_idx}"
            cell = ws.cell(row=row_idx, column=27, value=formula_pick_pack)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Упаковка (колонка 28 = AB)
            cell = ws.cell(row=row_idx, column=28, value=input_data.packaging_cost)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Эквайринг (колонка 29 = AC)
            formula_acquiring = f"=D{row_idx}*S{row_idx}/100"
            cell = ws.cell(row=row_idx, column=29, value=formula_acquiring)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Возвраты (колонка 30 = AD)
            formula_returns = f"=D{row_idx}*T{row_idx}/100"
            cell = ws.cell(row=row_idx, column=30, value=formula_returns)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Штрафы (колонка 31 = AE)
            penalty_prob = 0.35 if not input_data.has_night_shift else 0.05
            formula_penalty = f"=D{row_idx}*R{row_idx}/100*{penalty_prob}"
            cell = ws.cell(row=row_idx, column=31, value=formula_penalty)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Маркетинг (колонка 32 = AF)
            cell = ws.cell(row=row_idx, column=32, value=input_data.marketing_budget_per_unit)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Склад (колонка 33 = AG)
            formula_warehouse = f"=P{row_idx}/30/5"
            cell = ws.cell(row=row_idx, column=33, value=formula_warehouse)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Налог (колонка 34 = AH)
            formula_tax = f"=D{row_idx}*U{row_idx}/100"
            cell = ws.cell(row=row_idx, column=34, value=formula_tax)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # --- ИТОГИ ---
            # Итого расходов (колонка 35 = AI)
            formula_total = f"=E{row_idx}+X{row_idx}+Y{row_idx}+Z{row_idx}+AA{row_idx}+AB{row_idx}+AC{row_idx}+AD{row_idx}+AE{row_idx}+AF{row_idx}+AG{row_idx}+AH{row_idx}"
            cell = ws.cell(row=row_idx, column=35, value=formula_total)
            cell.fill = self.result_fill
            cell.font = Font(bold=True, size=11)
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Прибыль (колонка 36 = AJ)
            formula_profit = f"=D{row_idx}-AI{row_idx}"
            cell = ws.cell(row=row_idx, column=36, value=formula_profit)
            # Условное форматирование через формулу
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            if result.gross_profit > 0:
                cell.fill = self.profit_fill
            else:
                cell.fill = self.loss_fill
            
            # --- ДОПОЛНИТЕЛЬНЫЕ МЕТРИКИ ---
            # Маржа (колонка 37 = AK)
            formula_margin = f"=IF(D{row_idx}>0, (AJ{row_idx}/D{row_idx})*100, 0)"
            cell = ws.cell(row=row_idx, column=37, value=formula_margin)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '0.00"%"'
            
            # ROI (колонка 38 = AL)
            formula_roi = f"=IF(E{row_idx}>0, (AJ{row_idx}/E{row_idx})*100, 0)"
            cell = ws.cell(row=row_idx, column=38, value=formula_roi)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '0.00"%"'
            
            # Минимальная цена (колонка 39 = AM)
            formula_min_price = f"=MAX(0, (E{row_idx}+Y{row_idx}+Z{row_idx}+AA{row_idx}+AB{row_idx}+AF{row_idx}+AG{row_idx})/(1-Q{row_idx}/100-S{row_idx}/100-T{row_idx}/100-R{row_idx}/100*{penalty_prob}-U{row_idx}/100))"
            cell = ws.cell(row=row_idx, column=39, value=formula_min_price)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Максимальная скидка (колонка 40 = AN)
            formula_max_discount = f"=IF(D{row_idx}>0, ((D{row_idx}-AM{row_idx})/D{row_idx})*100, 0)"
            cell = ws.cell(row=row_idx, column=40, value=formula_max_discount)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '0.00"%"'
            
            # Точка безубыточности км (колонка 41 = AO)
            formula_break_even = f"=IF(Y{row_idx}>0, AJ{row_idx}/(K{row_idx}/L{row_idx}), 999999)"
            cell = ws.cell(row=row_idx, column=41, value=formula_break_even)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.0'
            
            # LTV (колонка 42 = AP)
            formula_ltv = f"=D{row_idx}*{input_data.avg_purchases_per_year}*{input_data.customer_retention_rate}/(1+{input_data.discount_rate})"
            cell = ws.cell(row=row_idx, column=42, value=formula_ltv)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # CAC (колонка 43 = AQ)
            formula_cac = f"=(AF{row_idx}+AE{row_idx}+Y{row_idx})/0.3"
            cell = ws.cell(row=row_idx, column=43, value=formula_cac)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # LTV/CAC (колонка 44 = AR)
            formula_ltv_cac = f"=IF(AQ{row_idx}>0, AP{row_idx}/AQ{row_idx}, 999)"
            cell = ws.cell(row=row_idx, column=44, value=formula_ltv_cac)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '0.0'
            
            # Прибыль FBO (колонка 45 = AS)
            fbo_mult = self.marketplace_config.fbo_multiplier
            formula_fbo_profit = f"=D{row_idx}-(E{row_idx}+X{row_idx}+Z{row_idx}*{fbo_mult}+AC{row_idx}+AD{row_idx}+AH{row_idx}+AF{row_idx}+AB{row_idx}+10)"
            cell = ws.cell(row=row_idx, column=45, value=formula_fbo_profit)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Прибыль FBP (колонка 46 = AT)
            fbp_mult = self.marketplace_config.fbp_multiplier
            formula_fbp_profit = f"=D{row_idx}-(E{row_idx}+X{row_idx}+Z{row_idx}*{fbp_mult}+AC{row_idx}+AD{row_idx}+AH{row_idx}+AF{row_idx}+AB{row_idx}+5)"
            cell = ws.cell(row=row_idx, column=46, value=formula_fbp_profit)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.number_format = '#,##0.00'
            
            # Рекомендуемая модель (колонка 47 = AU)
            formula_model = f'=IF(AND(AJ{row_idx}>=AS{row_idx}, AJ{row_idx}>=AT{row_idx}), "FBS", IF(AS{row_idx}>=AT{row_idx}, "FBO", "FBP"))'
            cell = ws.cell(row=row_idx, column=47, value=formula_model)
            cell.fill = self.result_fill
            cell.border = self.thin_border
            cell.font = Font(bold=True)
        
        # Добавление автофильтра
        last_col_letter = get_column_letter(len(all_headers))
        ws.auto_filter.ref = f"A3:{last_col_letter}{len(results) + 3}"
        
        # Условное форматирование для прибыли
        ws.conditional_formatting.add(
            f"AJ4:AJ{len(results) + 3}",
            CellIsRule(operator="greaterThan", formula=["0"], 
                      fill=PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid"),
                      font=Font(color="006100", bold=True))
        )
        ws.conditional_formatting.add(
            f"AJ4:AJ{len(results) + 3}",
            CellIsRule(operator="lessThan", formula=["0"], 
                      fill=PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"),
                      font=Font(color="9C0006", bold=True))
        )
    
    def _create_hidden_losses_sheet(self, ws, results, input_data_list, marketplace_name):
        """Создание листа со скрытыми потерями FBS"""
        
        ws.merge_cells('A1:G1')
        ws.cell(row=1, column=1, value=f"⚠️ Скрытые потери FBS - {marketplace_name}").font = self.title_font
        
        # Таблица скрытых потерь
        headers = [
            "Показатель FBS", "Формула расчета", "Значение (среднее)", 
            "Максимальное", "Минимальное", "Риск", "Точка контроля"
        ]
        
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.thin_border
        
        ws.column_dimensions['A'].width = 35
        ws.column_dimensions['B'].width = 45
        ws.column_dimensions['C'].width = 20
        ws.column_dimensions['D'].width = 20
        ws.column_dimensions['E'].width = 20
        ws.column_dimensions['F'].width = 30
        ws.column_dimensions['G'].width = 35
        
        # Расчет средних значений
        avg_first_mile = np.mean([r.first_mile_cost for r in results])
        avg_last_mile = np.mean([r.last_mile_cost for r in results])
        avg_pick_pack = np.mean([r.pick_pack_cost for r in results])
        avg_penalty = np.mean([r.penalty_cost for r in results])
        
        hidden_losses_data = [
            [
                "Себестоимость с доставкой до МП",
                "COGS + First Mile + Упаковка",
                f"{np.mean([d.cogs + r.first_mile_cost + r.packaging_cost for d, r in zip(input_data_list, results)]):.2f} ₽",
                f"{max([d.cogs + r.first_mile_cost + r.packaging_cost for d, r in zip(input_data_list, results)]):.2f} ₽",
                f"{min([d.cogs + r.first_mile_cost + r.packaging_cost for d, r in zip(input_data_list, results)]):.2f} ₽",
                "Высокий",
                "Контролировать вес брутто и расстояние"
            ],
            [
                "Полная логистика (до клиента)",
                "First Mile + Last Mile",
                f"{avg_first_mile + avg_last_mile:.2f} ₽",
                f"{max([r.first_mile_cost + r.last_mile_cost for r in results]):.2f} ₽",
                f"{min([r.first_mile_cost + r.last_mile_cost for r in results]):.2f} ₽",
                "Критический (>25% от цены)" if (avg_first_mile + avg_last_mile) / np.mean([d.selling_price for d in input_data_list]) > 0.25 else "Средний",
                "Оптимизация маршрутов и упаковки"
            ],
            [
                "Стоимость обработки заказа (Pick & Pack)",
                "Время сборки × Ставка оператора",
                f"{avg_pick_pack:.2f} ₽",
                f"{max([r.pick_pack_cost for r in results]):.2f} ₽",
                f"{min([r.pick_pack_cost for r in results]):.2f} ₽",
                "Средний",
                "Оптимизация процессов сборки"
            ],
            [
                "Штрафы за просрочку",
                "Цена × Ставка штрафа × Вероятность",
                f"{avg_penalty:.2f} ₽",
                f"{max([r.penalty_cost for r in results]):.2f} ₽",
                f"{min([r.penalty_cost for r in results]):.2f} ₽",
                "Высокий (при отсутствии ночной смены)",
                "Внедрить автоуведомления и ночную смену"
            ],
            [
                "Стоимость 1 часа просрочки",
                "(Цена × 5%) / Кол-во просрочек",
                f"{np.mean([d.selling_price * 0.05 for d in input_data_list]):.2f} ₽",
                "-",
                "-",
                "Критический",
                "Мониторинг времени обработки"
            ]
        ]
        
        for row_idx, row_data in enumerate(hidden_losses_data, 4):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self.thin_border
                if row_idx == 4:
                    cell.fill = PatternFill(start_color="FFE5E5", end_color="FFE5E5", fill_type="solid")
        
        # График структуры скрытых потерь
        chart = PieChart()
        chart.title = "Структура скрытых потерь FBS"
        chart.width = 20
        chart.height = 15
        
        data_ref = Reference(ws, min_col=3, min_row=4, max_row=8)
        cats_ref = Reference(ws, min_col=1, min_row=4, max_row=8)
        chart.add_data(data_ref, titles_from_data=False)
        chart.set_categories(cats_ref)
        
        ws.add_chart(chart, "A12")
    
    def _create_models_comparison_sheet(self, ws, results, input_data_list, marketplace_name):
        """Создание листа сравнения моделей FBS vs FBO vs FBP"""
        
        ws.merge_cells('A1:H1')
        ws.cell(row=1, column=1, 
                value=f"🔄 Сравнение моделей фулфилмента - {marketplace_name}").font = self.title_font
        
        headers = [
            "Артикул", "Наименование", "Прибыль FBS", "Прибыль FBO", 
            "Прибыль FBP", "Разница FBS-FBO", "Разница FBS-FBP", "Рекомендация"
        ]
        
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.thin_border
        
        for i, (result, input_data) in enumerate(zip(results, input_data_list), 4):
            ws.cell(row=i, column=1, value=result.artikul).border = self.thin_border
            ws.cell(row=i, column=2, value=result.product_name).border = self.thin_border
            ws.cell(row=i, column=3, value=result.gross_profit).border = self.thin_border
            ws.cell(row=i, column=4, value=result.fbo_profit).border = self.thin_border
            ws.cell(row=i, column=5, value=result.fbp_profit).border = self.thin_border
            ws.cell(row=i, column=6, value=result.gross_profit - result.fbo_profit).border = self.thin_border            ws.cell(row=i, column=7, value=result.gross_profit - result.fbp_profit).border = self.thin_border
            ws.cell(row=i, column=8, value=result.recommended_model).border = self.thin_border
            
            # Подсветка лучшей модели
            if result.recommended_model == "FBS":
                ws.cell(row=i, column=3).fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            elif result.recommended_model == "FBO":
                ws.cell(row=i, column=4).fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            else:
                ws.cell(row=i, column=5).fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    
    def _create_ltv_cac_sheet(self, ws, results, input_data_list, marketplace_name):
        """Создание листа с расчетом LTV и CAC"""
        
        ws.merge_cells('A1:I1')
        ws.cell(row=1, column=1, value=f"👥 Метрики LTV и CAC - {marketplace_name}").font = self.title_font
        
        headers = [
            "Артикул", "Средний чек", "Повторные покупки/год", "CRR", 
            "LTV", "CAC", "LTV/CAC", "Оценка", "Рекомендация"
        ]
        
        for col_idx, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.thin_border
        
        for i, (result, input_data) in enumerate(zip(results, input_data_list), 4):
            ws.cell(row=i, column=1, value=result.artikul).border = self.thin_border
            ws.cell(row=i, column=2, value=result.selling_price).border = self.thin_border
            ws.cell(row=i, column=3, value=input_data.avg_purchases_per_year).border = self.thin_border
            ws.cell(row=i, column=4, value=input_data.customer_retention_rate).border = self.thin_border
            ws.cell(row=i, column=5, value=result.ltv).border = self.thin_border
            ws.cell(row=i, column=6, value=result.cac).border = self.thin_border
            ws.cell(row=i, column=7, value=result.ltv_cac_ratio).border = self.thin_border
            
            # Оценка LTV/CAC
            if result.ltv_cac_ratio >= 3:
                assessment = "✅ Отлично"
                recommendation = "Масштабировать рекламу"
            elif result.ltv_cac_ratio >= 1:
                assessment = "⚠️ Нормально"
                recommendation = "Оптимизировать CAC"
            else:
                assessment = "❌ Плохо"
                recommendation = "Пересмотреть стратегию"
            
            ws.cell(row=i, column=8, value=assessment).border = self.thin_border
            ws.cell(row=i, column=9, value=recommendation).border = self.thin_border
            
            # Цветовое кодирование
            if result.ltv_cac_ratio >= 3:
                ws.cell(row=i, column=7).fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            elif result.ltv_cac_ratio >= 1:
                ws.cell(row=i, column=7).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
            else:
                ws.cell(row=i, column=7).fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    
    def _create_dashboard_sheet(self, ws, results, marketplace_name):
        """Создание дашборда с визуализацией"""
        
        ws.merge_cells('A1:L1')
        ws.cell(row=1, column=1, value=f"📈 Дашборд юнит-экономики - {marketplace_name}").font = self.title_font
        
        # Ключевые метрики
        total_profit = sum(r.gross_profit for r in results)
        avg_margin = np.mean([r.margin_percent for r in results])
        profitable_pct = len([r for r in results if r.gross_profit > 0]) / len(results) * 100
        
        metrics = [
            ("Общая прибыль", f"{total_profit:,.0f} ₽", "A4"),
            ("Средняя маржа", f"{avg_margin:.1f}%", "D4"),
            ("Прибыльных SKU", f"{profitable_pct:.1f}%", "G4"),
            ("Всего товаров", f"{len(results)}", "J4")
        ]
        
        for title, value, cell_ref in metrics:
            ws[cell_ref] = title
            ws[cell_ref].font = Font(bold=True, size=12)
            ws[f"{cell_ref[0]}{int(cell_ref[1:])+1}"] = value
            ws[f"{cell_ref[0]}{int(cell_ref[1:])+1}"].font = Font(size=14, color="1a1a2e", bold=True)
    
    def _create_recommendations_sheet(self, ws, results, marketplace_name):
        """Создание листа с рекомендациями"""
        
        ws.merge_cells('A1:D1')
        ws.cell(row=1, column=1, value=f"💡 Рекомендации по оптимизации - {marketplace_name}").font = self.title_font
        
        recommendations = []
        
        # Анализ убыточных товаров
        unprofitable = [r for r in results if r.gross_profit < 0]
        if unprofitable:
            recommendations.append([
                "Убыточные товары",
                f"{len(unprofitable)} из {len(results)}",
                "Высокий",
                "Пересмотреть цены, сменить поставщика или снять с продажи"
            ])
        
        # Анализ First Mile
        high_first_mile = [r for r in results if r.first_mile_cost > r.selling_price * 0.15]
        if high_first_mile:
            recommendations.append([
                "Высокая стоимость First Mile",
                f"{len(high_first_mile)} товаров",
                "Критический",
                "Оптимизировать логистику, увеличить загрузку паллет, рассмотреть FBO"
            ])
        
        # Анализ штрафов
        high_penalty = [r for r in results if r.penalty_cost > r.selling_price * 0.02]
        if high_penalty:
            recommendations.append([
                "Высокие штрафы за просрочку",
                f"{len(high_penalty)} товаров",
                "Высокий",
                "Внедрить ночную смену, ускорить обработку заказов"
            ])
        
        # Рекомендации по переходу на FBO
        fbo_better = [r for r in results if r.fbo_profit > r.gross_profit]
        if fbo_better:
            recommendations.append([
                "Переход на FBO выгоднее",
                f"{len(fbo_better)} товаров",
                "Средний",
                "Рассмотреть перевод части ассортимента на FBO"
            ])
        
        # Заголовки
        rec_headers = ["Проблема", "Масштаб", "Критичность", "Рекомендация"]
        for col_idx, header in enumerate(rec_headers, 1):
            cell = ws.cell(row=3, column=col_idx, value=header)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = self.header_alignment
            cell.border = self.thin_border
        
        # Запись рекомендаций
        for row_idx, rec in enumerate(recommendations, 4):
            for col_idx, value in enumerate(rec, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self.thin_border
    
    def _create_instructions_sheet(self, ws):
        """Создание листа с инструкцией"""
        
        ws.merge_cells('A1:C1')
        ws.cell(row=1, column=1, value="📖 Инструкция по использованию отчета").font = self.title_font
        
        instructions = [
            ["Раздел", "Описание", "Действия"],
            ["📊 Юнит-экономика FBS", "Основной расчет всех показателей", "Изменяйте желтые ячейки, зеленые пересчитаются автоматически"],
            ["⚠️ Скрытые потери FBS", "Анализ неочевидных расходов", "Обратите внимание на штрафы и Pick & Pack"],
            ["🔄 Сравнение моделей", "FBS vs FBO vs FBP", "Выберите оптимальную модель для каждого товара"],
            ["👥 LTV и CAC", "Метрики жизненной ценности клиента", "LTV/CAC должен быть > 3"],
            ["📈 Дашборд", "Ключевые показатели", "Быстрая оценка состояния бизнеса"],
            ["💡 Рекомендации", "Автоматические рекомендации", "Следуйте рекомендациям для оптимизации"]
        ]
        
        for row_idx, row_data in enumerate(instructions, 3):
            for col_idx, value in enumerate(row_data, 1):
                cell = ws.cell(row=row_idx, column=col_idx, value=value)
                cell.border = self.thin_border
                if row_idx == 3:
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="E8E8E8", end_color="E8E8E8", fill_type="solid")
        
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 40
        ws.column_dimensions['C'].width = 50

# ============================================================================
# БЛОК 7: ВИЗУАЛИЗАЦИЯ И ГРАФИКИ
# ============================================================================

class FBSVisualizer:
    """Класс для создания профессиональных визуализаций"""
    
    # Цветовая палитра
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
        'gradient_1': ['#00b894', '#00cec9', '#0984e3', '#6c5ce7', '#a29bfe'],
        'gradient_2': ['#d63031', '#e17055', '#fdcb6e', '#00b894', '#0984e3']
    }
    
    @staticmethod
    def create_cost_breakdown_pie(result: FBSResultData, title: str = "Структура расходов FBS") -> go.Figure:
        """Создание круговой диаграммы структуры расходов"""
        
        cost_categories = {
            'Себестоимость': result.total_expenses - sum([
                result.commission, result.first_mile_cost, result.last_mile_cost,
                result.pick_pack_cost, result.packaging_cost, result.acquiring_cost,
                result.return_cost, result.penalty_cost, result.marketing_cost,
                result.warehouse_cost, result.tax_cost
            ]),
            'Комиссия МП': result.commission,
            'First Mile': result.first_mile_cost,
            'Last Mile': result.last_mile_cost,
            'Pick & Pack': result.pick_pack_cost,
            'Упаковка': result.packaging_cost,
            'Эквайринг': result.acquiring_cost,
            'Возвраты': result.return_cost,
            'Штрафы': result.penalty_cost,
            'Маркетинг': result.marketing_cost,
            'Склад': result.warehouse_cost,
            'Налог': result.tax_cost
        }
        
        # Фильтруем нулевые значения
        cost_categories = {k: v for k, v in cost_categories.items() if v > 0.01}
        
        fig = go.Figure(data=[go.Pie(
            labels=list(cost_categories.keys()),
            values=list(cost_categories.values()),
            hole=0.4,
            marker=dict(colors=FBSVisualizer.COLORS['gradient_1'][:len(cost_categories)]),
            textinfo='label+percent',
            textposition='outside',
            textfont=dict(size=11),
            hovertemplate='<b>%{label}</b><br>Сумма: %{value:,.2f} ₽<br>Доля: %{percent}<extra></extra>'
        )])
        
        fig.update_layout(
            title=dict(
                text=f"<b>{title}</b><br><sub>Общие расходы: {result.total_expenses:,.2f} ₽</sub>",
                font=dict(size=16, color=FBSVisualizer.COLORS['primary']),
                x=0.5
            ),
            template="plotly_white",
            height=500,
            showlegend=True,
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=-0.2,
                xanchor="center",
                x=0.5
            ),
            margin=dict(t=100, b=100)
        )
        
        return fig
    
    @staticmethod
    def create_waterfall_chart(result: FBSResultData) -> go.Figure:
        """Создание водопадной диаграммы прибыли"""
        
        categories = [
            "Цена продажи", "Себестоимость", "Комиссия", "First Mile",
            "Last Mile", "Pick & Pack", "Упаковка", "Эквайринг",
            "Возвраты", "Штрафы", "Маркетинг", "Склад", "Налог", "ЧИСТАЯ ПРИБЫЛЬ"
        ]
        
        values = [
            result.selling_price,
            -(result.total_expenses - sum([
                result.commission, result.first_mile_cost, result.last_mile_cost,
                result.pick_pack_cost, result.packaging_cost, result.acquiring_cost,
                result.return_cost, result.penalty_cost, result.marketing_cost,
                result.warehouse_cost, result.tax_cost
            ])),
            -result.commission, -result.first_mile_cost, -result.last_mile_cost,
            -result.pick_pack_cost, -result.packaging_cost, -result.acquiring_cost,
            -result.return_cost, -result.penalty_cost, -result.marketing_cost,
            -result.warehouse_cost, -result.tax_cost,
            result.gross_profit
        ]
        
        # Цвета: зеленый для положительных, красный для отрицательных
        colors = [
            '#0984e3', '#d63031', '#d63031', '#e17055', '#e17055',
            '#fdcb6e', '#fdcb6e', '#d63031', '#d63031', '#e17055',
            '#fdcb6e', '#fdcb6e', '#d63031', '#00b894' if result.gross_profit > 0 else '#d63031'
        ]
        
        fig = go.Figure(data=[go.Waterfall(
            name="Прибыль",
            orientation="v",
            measure=["absolute"] + ["relative"] * 12 + ["total"],
            x=categories,
            y=values,
            text=[f"{v:,.0f} ₽" for v in values],
            textposition="outside",
            connector={"line": {"color": "rgb(63, 63, 63)"}},
            increasing={"marker": {"color": "#00b894"}},
            decreasing={"marker": {"color": "#d63031"}},
            totals={"marker": {"color": "#0984e3" if result.gross_profit > 0 else "#d63031"}}
        )])
        
        fig.update_layout(
            title=dict(
                text="<b>Водопадная диаграмма формирования прибыли</b>",
                font=dict(size=16, color=FBSVisualizer.COLORS['primary'])
            ),
            template="plotly_white",
            height=500,
            showlegend=False,
            xaxis=dict(tickangle=45),
            margin=dict(t=80, b=100)
        )
        
        return fig
    
    @staticmethod
    def create_models_comparison_chart(results: List[FBSResultData]) -> go.Figure:
        """Создание сравнительной диаграммы моделей FBS/FBO/FBP"""
        
        if not results:
            return go.Figure()
        
        # Топ-10 товаров по прибыли
        top_results = sorted(results, key=lambda x: x.gross_profit, reverse=True)[:10]
        
        artikuls = [r.artikul[:15] for r in top_results]
        fbs_profits = [r.gross_profit for r in top_results]
        fbo_profits = [r.fbo_profit for r in top_results]
        fbp_profits = [r.fbp_profit for r in top_results]
        
        fig = go.Figure()
        
        fig.add_trace(go.Bar(
            name='FBS',
            x=artikuls,
            y=fbs_profits,
            marker_color='#0984e3',
            text=[f'{v:,.0f}' for v in fbs_profits],
            textposition='auto'
        ))
        
        fig.add_trace(go.Bar(
            name='FBO',
            x=artikuls,
            y=fbo_profits,
            marker_color='#00b894',
            text=[f'{v:,.0f}' for v in fbo_profits],
            textposition='auto'
        ))
        
        fig.add_trace(go.Bar(
            name='FBP',
            x=artikuls,
            y=fbp_profits,
            marker_color='#6c5ce7',
            text=[f'{v:,.0f}' for v in fbp_profits],
            textposition='auto'
        ))
        
        fig.update_layout(
            title=dict(
                text="<b>Сравнение моделей: FBS vs FBO vs FBP</b>",
                font=dict(size=16)
            ),
            barmode='group',
            template="plotly_white",
            height=450,
            xaxis=dict(tickangle=45),
            legend=dict(
                orientation="h",
                yanchor="bottom",
                y=1.02,
                xanchor="right",
                x=1
            )
        )
        
        return fig
    
    @staticmethod
    def create_break_even_distance_chart(results: List[FBSResultData]) -> go.Figure:
        """Визуализация точки безубыточности по расстоянию"""
        
        if not results:
            return go.Figure()
        
        # Топ товаров по точке безубыточности
        valid_results = [r for r in results if r.break_even_distance_km < 999999]
        valid_results = sorted(valid_results, key=lambda x: x.break_even_distance_km)[:15]
        
        if not valid_results:
            return go.Figure()
        
        fig = go.Figure(data=[go.Bar(
            x=[r.artikul[:15] for r in valid_results],
            y=[r.break_even_distance_km for r in valid_results],
            marker=dict(
                color=[r.break_even_distance_km for r in valid_results],
                colorscale='RdYlGn',
                showscale=True,
                colorbar=dict(title="Км")
            ),
            text=[f'{r.break_even_distance_km:.1f} км' for r in valid_results],
            textposition='auto'
        )])
        
        fig.add_hline(
            y=50, line_dash="dash", line_color="red",
            annotation_text="Критическая зона (< 50 км)",
            annotation_position="bottom right"
        )
        
        fig.update_layout(
            title="<b>Точка безубыточности по расстоянию (First Mile)</b>",
            xaxis_title="Товар",
            yaxis_title="Максимальное расстояние, км",
            template="plotly_white",
            height=450,
            xaxis=dict(tickangle=45)
        )
        
        return fig

# ============================================================================
# БЛОК 8: ИНТЕРФЕЙС ПОЛЬЗОВАТЕЛЯ (STREAMLIT)
# ============================================================================

def init_session_state():
    """Инициализация всех состояний сессии"""
    
    if 'calculator' not in st.session_state:
        st.session_state.calculator = FBSUnitEconomicsCalculator()
    
    if 'visualizer' not in st.session_state:
        st.session_state.visualizer = FBSVisualizer()
    
    if 'secure_data' not in st.session_state:
        st.session_state.secure_data = SecureDataManager()
    
    if 'cache_manager' not in st.session_state:
        st.session_state.cache_manager = CacheManager()
    
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
    
    if 'calculation_history' not in st.session_state:
        st.session_state.calculation_history = []

def render_sidebar():
    """Отрисовка боковой панели навигации"""
    
    with st.sidebar:
        st.markdown("""
        <div style='text-align: center; padding: 15px; background: linear-gradient(135deg, #1a1a2e, #16213e); border-radius: 10px; margin-bottom: 20px;'>
            <h2 style='color: white; margin: 0;'>🚀 FBS PRO</h2>
            <p style='color: #a8a8d0; margin: 5px 0;'>Юнит-экономика 2026</p>
        </div>
        """, unsafe_allow_html=True)
        
        # Навигация
        st.markdown("### 🧭 Навигация")
        
        sections = {
            "🏠 Главная": "main",
            "🧮 Калькулятор FBS": "calculator",
            "📊 Массовый расчет": "batch",
            "📈 Дашборд": "dashboard",
            "📥 Экспорт": "export",
            "⚙️ Настройки": "settings",
            "📖 Справка": "help"
        }
        
        selected_section = st.radio(
            "Выберите раздел:",
            list(sections.keys()),
            label_visibility="collapsed"
        )
        
        st.session_state.current_section = sections[selected_section]
        
        # Статус системы
        st.markdown("---")
        st.markdown("### 📊 Статус")
        
        col1, col2 = st.columns(2)
        with col1:
            st.metric("Маркетплейс", st.session_state.marketplace)
        with col2:
            st.metric("Налоги", st.session_state.tax_system.split()[0])
        
        if st.session_state.results:
            st.success(f"✅ Рассчитано: {len(st.session_state.results)} товаров")
        else:
            st.info("ℹ️ Нет расчетов")
        
        # Быстрые действия
        st.markdown("---")
        st.markdown("### ⚡ Быстрые действия")
        
        if st.button("🧹 Очистить кэш", use_container_width=True):
            st.session_state.cache_manager.clear_cache()
            st.success("Кэш очищен!")
        
        if st.button("📋 История расчетов", use_container_width=True):
            st.session_state.show_history = not st.session_state.get('show_history', False)
        
        # Информация
        st.markdown("---")
        st.markdown(f"""
        <div style='text-align: center; color: #666; font-size: 0.8em;'>
            <p>Версия {APP_VERSION}</p>
            <p>© 2026 FBS PRO</p>
        </div>
        """, unsafe_allow_html=True)

def render_main_page():
    """Главная страница"""
    
    st.markdown("""
    <div style='text-align: center; padding: 40px 20px; background: linear-gradient(135deg, #1a1a2e, #16213e, #0f3460); border-radius: 15px; margin-bottom: 30px;'>
        <h1 style='color: white; font-size: 2.5em; margin: 0;'>🚀 FBS Юнит-экономика PRO</h1>
        <p style='color: #a8a8d0; font-size: 1.2em; margin: 15px 0;'>
            Профессиональный расчет юнит-экономики для FBS-модели
        </p>
        <p style='color: #6666aa; font-size: 0.9em;'>
            Ozon • Wildberries • Яндекс Маркет | First Mile + Last Mile | Pick & Pack | Штрафы
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Ключевые возможности
    st.markdown("### 🎯 Ключевые возможности")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #0984e3, #6c5ce7); padding: 20px; border-radius: 10px; color: white;'>
            <h4>📦 Двойная логистика</h4>
            <p>Расчет First Mile (ваша доставка до МП) + Last Mile (доставка МП клиенту)</p>
            <p><b>Точка безубыточности по расстоянию</b></p>
        </div>
        """, unsafe_allow_html=True)
    
    with col2:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #00b894, #00cec9); padding: 20px; border-radius: 10px; color: white;'>
            <h4>⚠️ Скрытые потери</h4>
            <p>Штрафы за просрочку, Pick & Pack, стоимость обработки заказа</p>
            <p><b>Вероятность просрочки и оптимизация</b></p>
        </div>
        """, unsafe_allow_html=True)
    
    with col3:
        st.markdown("""
        <div style='background: linear-gradient(135deg, #e17055, #d63031); padding: 20px; border-radius: 10px; color: white;'>
            <h4>📊 LTV и CAC</h4>
            <p>Расчет жизненной ценности клиента и стоимости привлечения</p>
            <p><b>Сравнение моделей FBS/FBO/FBP</b></p>
        </div>
        """, unsafe_allow_html=True)
    
    # Быстрый старт
    st.markdown("---")
    st.markdown("### 🚀 Быстрый старт")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.info("**Шаг 1:** Настройте маркетплейс и налоговую систему в разделе ⚙️ Настройки")
    
    with col2:
        st.info("**Шаг 2:** Выполните расчет в разделе 🧮 Калькулятор FBS")
    
    with col3:
        st.info("**Шаг 3:** Экспортируйте результаты в Excel с живыми формулами")
    
    # Статистика (если есть расчеты)
    if st.session_state.results:
        st.markdown("---")
        st.markdown("### 📈 Последние результаты")
        
        results = st.session_state.results
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.metric("Товаров", len(results))
        with col2:
            avg_margin = np.mean([r.margin_percent for r in results])
            st.metric("Средняя маржа", f"{avg_margin:.1f}%")
        with col3:
            profitable = len([r for r in results if r.gross_profit > 0])
            st.metric("Прибыльных", f"{profitable}/{len(results)}")
        with col4:
            total_profit = sum(r.gross_profit for r in results)
            st.metric("Общая прибыль", f"{total_profit:,.0f} ₽")

def render_calculator_page():
    """Страница калькулятора FBS"""
    
    st.markdown("## 🧮 Калькулятор FBS юнит-экономики")
    
    st.info("""
    **🎯 Этот калькулятор учитывает специфику FBS:**
    - Двойную логистику (First Mile + Last Mile)
    - Штрафы за просрочку передачи заказа
    - Стоимость обработки заказа (Pick & Pack)
    - Точку безубыточности по расстоянию
    - Запас прочности для сезонных распродаж
    """)
    
    # Выбор режима
    calc_mode = st.radio(
        "Режим расчета:",
        ["📱 Один товар", "📊 Массовый расчет"],
        horizontal=True
    )
    
    if calc_mode == "📱 Один товар":
        render_single_calculator()
    else:
        render_batch_calculator()

def render_single_calculator():
    """Калькулятор для одного товара"""
    
    # Создаем две колонки для ввода данных
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("📦 Основные параметры")
        
        artikul = st.text_input("Артикул", value="SKU-001", key="single_artikul")
        product_name = st.text_input("Наименование товара", value="Тестовый товар", key="single_name")
        category = st.selectbox(
            "Категория",
            options=["default", "auto_parts", "electronics", "clothing", "home", "sport", "beauty", "books", "toys", "food"],
            format_func=lambda x: {
                "default": "Общая категория",
                "auto_parts": "Автозапчасти",
                "electronics": "Электроника",
                "clothing": "Одежда",
                "home": "Товары для дома",
                "sport": "Спорт",
                "beauty": "Красота",
                "books": "Книги",
                "toys": "Игрушки",
                "food": "Продукты"
            }[x],
            key="single_category"
        )
        
        st.markdown("---")
        st.markdown("**💰 Финансы**")
        selling_price = st.number_input("Цена продажи, ₽", value=5000.0, step=100.0, key="single_price")
        cogs = st.number_input("Себестоимость закупки, ₽", value=3000.0, step=100.0, key="single_cogs")
        
        st.markdown("---")
        st.markdown("**📏 Габариты**")
        weight = st.number_input("Вес брутто, кг", value=1.5, step=0.1, key="single_weight")
        col_dim1, col_dim2, col_dim3 = st.columns(3)
        with col_dim1:
            length = st.number_input("Длина, см", value=20, key="single_length")
        with col_dim2:
            width = st.number_input("Ширина, см", value=15, key="single_width")
        with col_dim3:
            height = st.number_input("Высота, см", value=10, key="single_height")
    
    with col2:
        st.subheader("🚚 FBS специфика")
        
        st.markdown("**🚛 First Mile (Ваша логистика до МП)**")
        warehouse_distance = st.number_input("Расстояние до склада МП, км", value=50.0, step=1.0, key="single_distance")
        transport_cost_per_km = st.number_input("Стоимость 1 км транспорта, ₽", value=20.0, step=1.0, key="single_km_cost")
        pallet_capacity = st.number_input("Единиц на паллете", value=100, step=10, key="single_pallet")
        
        st.markdown("---")
        st.markdown("**📦 Обработка заказа**")
        pick_pack_time = st.number_input("Время сборки заказа, мин", value=5.0, step=0.5, key="single_pick_time")
        operator_rate = st.number_input("Ставка оператора сборки, ₽/час", value=300.0, step=50.0, key="single_operator_rate")
        packaging_cost = st.number_input("Стоимость упаковки, ₽/шт", value=50.0, step=10.0, key="single_packaging")
        
        st.markdown("---")
        st.markdown("**⚠️ Риски**")
        has_night_shift = st.checkbox("Наличие ночной смены", value=False, key="single_night_shift",
                                       help="Снижает вероятность штрафов за просрочку с 35% до 5%")
        
        st.markdown("---")
        st.markdown("**📊 Маркетинг и клиенты**")
        marketing_budget = st.number_input("Маркетинговый бюджет на ед., ₽", value=100.0, step=10.0, key="single_marketing")
        avg_purchases = st.number_input("Среднее кол-во покупок в год", value=2.5, step=0.1, key="single_purchases")
        crr = st.number_input("Коэффициент удержания (CRR)", value=0.7, step=0.05, min_value=0.0, max_value=1.0, key="single_crr")
    
    # Кнопка расчета
    st.markdown("---")
    if st.button("🚀 Рассчитать юнит-экономику", type="primary", use_container_width=True, key="single_calc_btn"):
        with st.spinner("Выполняется профессиональный расчет FBS..."):
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
                customer_retention_rate=crr
            )
            
            # Выполнение расчета
            calculator = st.session_state.calculator
            calculator.set_marketplace(st.session_state.marketplace)
            calculator.tax_system = st.session_state.tax_system
            
            result = calculator.calculate_unit_economics(input_data)
            
            # Сохранение в историю
            st.session_state.calculation_history.append({
                'timestamp': datetime.now(),
                'input': input_data,
                'result': result
            })
            
            # Отображение результатов
            render_single_result(result, input_data)

def render_single_result(result: FBSResultData, input_data: FBSInputData):
    """Отображение результатов расчета одного товара"""
    
    st.markdown("---")
    st.markdown("## 📊 Результаты расчета FBS")
    
    # Ключевые метрики
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        profit_color = "green" if result.gross_profit > 0 else "red"
        st.metric(
            "💰 Чистая прибыль",
            f"{result.gross_profit:,.2f} ₽",
            delta=f"{result.margin_percent:.1f}% маржи"
        )
    
    with col2:
        st.metric("📦 Общие расходы", f"{result.total_expenses:,.2f} ₽")
    
    with col3:
        st.metric("📈 ROI", f"{result.roi_percent:.1f}%")
    
    with col4:
        st.metric("💡 Рек. модель", result.recommended_model)
    
    # Детализация расходов
    st.markdown("### 📋 Детализация расходов FBS")
    
    # Создаем DataFrame для отображения
    expenses_data = {
        "Статья расходов": [
            "Себестоимость закупки",
            "Комиссия маркетплейса",
            "🚛 First Mile (доставка до МП)",
            "📦 Last Mile (доставка клиенту)",
            "👷 Pick & Pack (обработка заказа)",
            "📦 Упаковка",
            "💳 Эквайринг",
            "↩️ Возвраты",
            "⚠️ Штрафы за просрочку",
            "📊 Маркетинг",
            "🏭 Складские расходы",
            "💰 Налог",
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
            f"{input_data.cogs / result.selling_price * 100:.1f}%" if result.selling_price > 0 else "0%",
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
    st.dataframe(df_expenses, use_container_width=True, height=500)
    
    # FBS специфические метрики
    st.markdown("### ⚠️ Специфические метрики FBS")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("Вероятность просрочки", f"{result.penalty_probability:.0%}")
        st.metric("Штрафы за просрочку", f"{result.penalty_cost:.2f} ₽")
    
    with col2:
        st.metric("Точка безубыточности (км)", f"{result.break_even_distance_km:.1f} км")
        if result.break_even_distance_km < 50:
            st.warning("⚠️ Критически близко! Риск убытков при увеличении расстояния")
    
    with col3:
        st.metric("Максимальная скидка", f"{result.max_discount_percent:.1f}%")
        st.metric("Запас прочности по цене", f"{result.safety_margin_price:.2f} ₽")
    
    # LTV и CAC
    st.markdown("### 👥 LTV и CAC")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("LTV", f"{result.ltv:,.2f} ₽")
    with col2:
        st.metric("CAC", f"{result.cac:,.2f} ₽")
    with col3:
        ltv_cac_color = "green" if result.ltv_cac_ratio >= 3 else ("orange" if result.ltv_cac_ratio >= 1 else "red")
        st.metric("LTV/CAC", f"{result.ltv_cac_ratio:.1f}")
        if result.ltv_cac_ratio >= 3:
            st.success("✅ Отличное соотношение!")
        elif result.ltv_cac_ratio >= 1:
            st.warning("⚠️ Приемлемо, но можно улучшить")
        else:
            st.error("❌ Плохое соотношение, пересмотрите стратегию")
    
    # Сравнение моделей
    st.markdown("### 🔄 Сравнение с другими моделями")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric("FBS", f"{result.gross_profit:,.2f} ₽",
                 delta="Текущая модель" if result.recommended_model == "FBS" else "")
    with col2:
        st.metric("FBO", f"{result.fbo_profit:,.2f} ₽",
                 delta="Рекомендуется" if result.recommended_model == "FBO" else "")
    with col3:
        st.metric("FBP", f"{result.fbp_profit:,.2f} ₽",
                 delta="Рекомендуется" if result.recommended_model == "FBP" else "")
    
    # Визуализация
    st.markdown("---")
    st.markdown("## 📊 Визуализация")
    
    col1, col2 = st.columns(2)
    
    with col1:
        fig_pie = st.session_state.visualizer.create_cost_breakdown_pie(result)
        st.plotly_chart(fig_pie, use_container_width=True)
    
    with col2:
        fig_waterfall = st.session_state.visualizer.create_waterfall_chart(result)
        st.plotly_chart(fig_waterfall, use_container_width=True)
    
    # Рекомендации
    st.markdown("---")
    st.markdown("## 💡 Рекомендации")
    
    recommendations = []
    
    if result.gross_profit <= 0:
        recommendations.append("❌ **Товар убыточен!** Пересмотрите цену или найдите поставщика с более низкой себестоимостью.")
    
    if result.first_mile_cost > result.selling_price * 0.15:
        recommendations.append("⚠️ **Высокая стоимость First Mile!** Рассмотрите переход на FBO или оптимизируйте логистику.")
    
    if result.penalty_cost > result.selling_price * 0.02:
        recommendations.append("⚠️ **Высокие штрафы за просрочку!** Внедрите ночную смену или ускорьте обработку заказов.")
    
    if result.recommended_model != "FBS":
        recommendations.append(f"💡 **Модель {result.recommended_model} выгоднее!** Рассмотрите переход с FBS на {result.recommended_model}.")
    
    if result.ltv_cac_ratio < 1:
        recommendations.append("❌ **LTV/CAC < 1!** Вы тратите на привлечение клиента больше, чем он приносит.")
    
    if result.max_discount_percent < 15:
        recommendations.append("⚠️ **Малый запас для скидок!** Вы не сможете участвовать в крупных распродажах без убытка.")
    
    if not recommendations:
        recommendations.append("✅ **Отличные показатели!** Товар прибыльный, продолжайте в том же духе!")
    
    for rec in recommendations:
        st.markdown(rec)

def render_batch_calculator():
    """Массовый расчет"""
    
    st.subheader("📊 Массовый расчет FBS")
    
    st.info("""
    **Загрузите файл с данными о товарах для массового расчета.**
    
    **Обязательные колонки:**
    - Артикул (или аналогичное название)
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
    
    # Загрузка файла
    uploaded_file = st.file_uploader(
        "📁 Загрузите файл каталога (CSV или Excel)",
        type=['csv', 'xlsx', 'xls'],
        help="Файл должен содержать колонки с артикулами, ценами и себестоимостью"
    )
    
    if uploaded_file is not None:
        try:
            # Чтение файла
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, encoding='utf-8-sig', dtype=str)
            else:
                df = pd.read_excel(uploaded_file, dtype=str)
            
            st.success(f"✅ Загружено {len(df)} товаров")
            st.dataframe(df.head(10), use_container_width=True)
            
            # Настройка маппинга колонок
            st.markdown("### 🔧 Настройка маппинга колонок")
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                artikul_col = st.selectbox("Колонка с артикулом", df.columns, 
                                           index=next((i for i, c in enumerate(df.columns) if 'артикул' in c.lower()), 0))
            with col2:
                price_col = st.selectbox("Колонка с ценой", df.columns,
                                        index=next((i for i, c in enumerate(df.columns) if 'цен' in c.lower()), 0))
            with col3:
                cost_col = st.selectbox("Колонка с себестоимостью", df.columns,
                                       index=next((i for i, c in enumerate(df.columns) if 'себестоимость' in c.lower() or 'закуп' in c.lower()), 0))
            with col4:
                name_col = st.selectbox("Колонка с наименованием", ["Не выбрано"] + list(df.columns), index=0)
            
            col1, col2, col3, col4 = st.columns(4)
            
            with col1:
                weight_col = st.selectbox("Колонка с весом (кг)", ["Не выбрано"] + list(df.columns),
                                         index=next((i+1 for i, c in enumerate(df.columns) if 'вес' in c.lower()), 0))
            with col2:
                length_col = st.selectbox("Колонка с длиной (см)", ["Не выбрано"] + list(df.columns),
                                         index=next((i+1 for i, c in enumerate(df.columns) if 'длин' in c.lower()), 0))
            with col3:
                width_col = st.selectbox("Колонка с шириной (см)", ["Не выбрано"] + list(df.columns),
                                        index=next((i+1 for i, c in enumerate(df.columns) if 'ширин' in c.lower()), 0))
            with col4:
                height_col = st.selectbox("Колонка с высотой (см)", ["Не выбрано"] + list(df.columns),
                                         index=next((i+1 for i, c in enumerate(df.columns) if 'высот' in c.lower()), 0))
            
            # Общие параметры
            st.markdown("### ⚙️ Общие параметры расчета")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                default_distance = st.number_input("Расстояние до МП (км)", value=50.0, step=1.0)
                default_transport_cost = st.number_input("Стоимость 1 км (₽)", value=20.0, step=1.0)
            
            with col2:
                default_pallet = st.number_input("Единиц на паллете", value=100, step=10)
                default_packaging = st.number_input("Упаковка (₽/шт)", value=50.0, step=10.0)
            
            with col3:
                default_pick_time = st.number_input("Время сборки (мин)", value=5.0, step=0.5)
                default_operator_rate = st.number_input("Ставка оператора (₽/ч)", value=300.0, step=50.0)
            
            # Кнопка расчета
            if st.button("🚀 Выполнить массовый расчет", type="primary", use_container_width=True):
                with st.spinner(f"Расчет {len(df)} товаров..."):
                    # Создание входных данных
                    input_data_list = []
                    
                    progress_bar = st.progress(0)
                    
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
                                marketing_budget_per_unit=100.0
                            )
                            input_data_list.append(input_data)
                        except Exception as e:
                            st.warning(f"⚠️ Ошибка в строке {idx}: {e}")
                            continue
                        
                        if idx % 100 == 0:
                            progress_bar.progress(min(idx / len(df), 1.0))
                    
                    # Выполнение расчета
                    calculator = st.session_state.calculator
                    calculator.set_marketplace(st.session_state.marketplace)
                    calculator.tax_system = st.session_state.tax_system
                    
                    results = calculator.calculate_batch(input_data_list)
                    
                    # Сохранение результатов
                    st.session_state.results = results
                    st.session_state.input_data_list = input_data_list
                    
                    progress_bar.progress(1.0)
                    st.success(f"✅ Рассчитано {len(results)} товаров!")
                    
                    # Краткая статистика
                    if results:
                        col1, col2, col3 = st.columns(3)
                        
                        with col1:
                            avg_margin = np.mean([r.margin_percent for r in results])
                            st.metric("Средняя маржа", f"{avg_margin:.1f}%")
                        
                        with col2:
                            profitable = len([r for r in results if r.gross_profit > 0])
                            st.metric("Прибыльных", f"{profitable}/{len(results)}")
                        
                        with col3:
                            total_profit = sum(r.gross_profit for r in results)
                            st.metric("Общая прибыль", f"{total_profit:,.0f} ₽")
        except Exception as e:
            st.error(f"❌ Ошибка чтения файла: {e}")
            logger.exception("Ошибка в batch calculator")

def render_dashboard_page():
    """Страница дашборда"""
    
    st.markdown("## 📈 Дашборд юнит-экономики")
    
    if not st.session_state.results:
        st.warning("⚠️ Нет данных для отображения. Выполните расчет в разделе 'Калькулятор FBS'.")
        return
    
    results = st.session_state.results
    
    # Ключевые метрики
    col1, col2, col3, col4, col5 = st.columns(5)
    
    with col1:
        st.metric("Товаров", len(results))
    with col2:
        avg_margin = np.mean([r.margin_percent for r in results])
        st.metric("Средняя маржа", f"{avg_margin:.1f}%")
    with col3:
        total_profit = sum(r.gross_profit for r in results)
        st.metric("Общая прибыль", f"{total_profit:,.0f} ₽")
    with col4:
        profitable = len([r for r in results if r.gross_profit > 0])
        st.metric("Прибыльных", f"{profitable}")
    with col5:
        fbo_better = len([r for r in results if r.fbo_profit > r.gross_profit])
        st.metric("FBO выгоднее", f"{fbo_better}")
    
    # Графики
    st.markdown("---")
    
    col1, col2 = st.columns(2)
    
    with col1:
        # Сравнение моделей
        fig_models = st.session_state.visualizer.create_models_comparison_chart(results)
        st.plotly_chart(fig_models, use_container_width=True)
    
    with col2:
        # Точка безубыточности
        fig_break_even = st.session_state.visualizer.create_break_even_distance_chart(results)
        st.plotly_chart(fig_break_even, use_container_width=True)
    
    # Таблица результатов
    st.markdown("---")
    st.markdown("### 📋 Детальные результаты")
    
    # Конвертация в DataFrame
    df_results = pd.DataFrame([r.to_dict() for r in results])
    
    # Выбор колонок для отображения
    display_cols = [
        'artikul', 'product_name', 'selling_price', 'total_expenses',
        'gross_profit', 'margin_percent', 'roi_percent',
        'first_mile_cost', 'last_mile_cost', 'penalty_cost',
        'ltv', 'cac', 'recommended_model'
    ]
    
    available_cols = [c for c in display_cols if c in df_results.columns]
    st.dataframe(df_results[available_cols], use_container_width=True, height=400)

def render_export_page():
    """Страница экспорта"""
    
    st.markdown("## 📥 Экспорт результатов")
    
    if not st.session_state.results:
        st.warning("⚠️ Нет данных для экспорта. Выполните расчет в разделе 'Калькулятор FBS'.")
        return
    
    results = st.session_state.results
    input_data_list = st.session_state.input_data_list
    
    st.success(f"✅ Доступно для экспорта: {len(results)} товаров")
    
    # Опции экспорта
    st.markdown("### 📊 Экспорт в Excel с живыми формулами")
    
    st.info("""
    **Профессиональный Excel-отчет включает:**
    - 📊 Основной лист с юнит-экономикой и живыми формулами
    - ⚠️ Лист со скрытыми потерями FBS
    - 🔄 Сравнение моделей FBS/FBO/FBP
    - 👥 Расчет LTV и CAC
    - 📈 Дашборд с визуализацией
    - 💡 Автоматические рекомендации
    - 📖 Инструкцию по использованию
    """)
    
    if st.button("📥 Скачать Excel отчет", type="primary", use_container_width=True):
        if st.session_state.exporter is None:
            st.error("❌ OpenPyXL не установлен. Выполните: pip install openpyxl")
        else:
            with st.spinner("Создание профессионального Excel-отчета..."):
                try:
                    output_path = EXPORTS_DIR / f"FBS_unit_economics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    
                    success = st.session_state.exporter.export_fbs_report(
                        results=results,
                        input_data_list=input_data_list,
                        marketplace_name=st.session_state.marketplace,
                        output_path=str(output_path)
                    )
                    
                    if success and output_path.exists():
                        with open(output_path, "rb") as f:
                            file_data = f.read()
                        
                        st.download_button(
                            label="⬇️ Скачать Excel отчет",
                            data=file_data,
                            file_name=output_path.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="download_excel"
                        )
                        st.success("✅ Excel отчет готов к скачиванию!")
                    else:
                        st.error("❌ Ошибка создания отчета")
                except Exception as e:
                    st.error(f"❌ Ошибка: {e}")
                    logger.exception("Ошибка экспорта")
    
    # Экспорт в CSV
    st.markdown("---")
    st.markdown("### 📄 Экспорт в CSV")
    
    if st.button("📄 Скачать CSV", use_container_width=True):
        df_results = pd.DataFrame([r.to_dict() for r in results])
        csv_data = df_results.to_csv(index=False, encoding='utf-8-sig')
        
        st.download_button(
            label="⬇️ Скачать CSV файл",
            data=csv_data,
            file_name=f"FBS_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv",
            key="download_csv"
        )

def render_settings_page():
    """Страница настроек"""
    
    st.markdown("## ⚙️ Настройки")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🏪 Маркетплейс")
        
        marketplace = st.selectbox(
            "Выберите маркетплейс",
            options=["Ozon", "Wildberries", "Яндекс Маркет"],
            index=list(MARKETPLACE_CONFIGS.keys()).index(st.session_state.marketplace) 
                  if st.session_state.marketplace in MARKETPLACE_CONFIGS else 0
        )
        
        if st.button("💾 Сохранить маркетплейс"):
            st.session_state.marketplace = marketplace
            st.session_state.calculator.set_marketplace(marketplace)
            st.success(f"✅ Маркетплейс '{marketplace}' сохранен!")
        
        # Отображение текущих тарифов
        config = MARKETPLACE_CONFIGS[marketplace]
        st.markdown("**Текущие тарифы:**")
        
        tariffs_df = pd.DataFrame({
            "Параметр": [
                "Базовая комиссия",
                "Мин. комиссия",
                "Last Mile (база)",
                "Last Mile (за кг)",
                "Эквайринг",
                "Возвраты",
                "Штраф за просрочку",
                "Время на передачу",
                "FBO множитель",
                "FBP множитель"
            ],
            "Значение": [
                f"{config.commission_rates['default']:.1%}",
                f"{config.min_commission} ₽",
                f"{config.last_mile_base} ₽",
                f"{config.last_mile_per_kg} ₽/кг",
                f"{config.acquiring_fee:.1%}",
                f"{config.return_fee:.1%}",
                f"{config.penalty_rate:.1%}",
                f"{config.penalty_time_hours} ч",
                f"{config.fbo_multiplier:.2f}",
                f"{config.fbp_multiplier:.2f}"
            ]
        })
        
        st.dataframe(tariffs_df, use_container_width=True)
    
    with col2:
        st.subheader("💰 Налоговая система")
        
        tax_system = st.selectbox(
            "Выберите систему налогообложения",
            options=list(TAX_SYSTEMS.keys()),
            index=list(TAX_SYSTEMS.keys()).index(st.session_state.tax_system)
                  if st.session_state.tax_system in TAX_SYSTEMS else 0
        )
        
        if st.button("💾 Сохранить налоговую систему"):
            st.session_state.tax_system = tax_system
            st.session_state.calculator.tax_system = tax_system
            st.success(f"✅ Налоговая система '{tax_system}' сохранена!")
        
        # Информация о налоговой системе
        tax_config = TAX_SYSTEMS[tax_system]
        st.markdown(f"""
        **Информация о налоговой системе:**
        - **Ставка:** {tax_config['rate']:.1%}
        - **База:** {'Доходы' if tax_config['base'] == 'revenue' else 'Прибыль'}
        - **Тип:** {tax_config['name']}
        """)
    
    # Дополнительные настройки
    st.markdown("---")
    st.subheader("🔧 Дополнительные настройки")
    
    col1, col2 = st.columns(2)
    
    with col1:
        st.markdown("**Настройки по умолчанию для расчета:**")
        
        default_distance = st.number_input("Расстояние до МП по умолчанию (км)", value=50.0, step=1.0)
        default_transport_cost = st.number_input("Стоимость 1 км по умолчанию (₽)", value=20.0, step=1.0)
        default_pallet = st.number_input("Единиц на паллете по умолчанию", value=100, step=10)
        
        if st.button("💾 Сохранить настройки по умолчанию"):
            st.session_state.default_distance = default_distance
            st.session_state.default_transport_cost = default_transport_cost
            st.session_state.default_pallet = default_pallet
            st.success("✅ Настройки сохранены!")
    
    with col2:
        st.markdown("**Управление данными:**")
        
        if st.button("🗑️ Очистить все расчеты"):
            st.session_state.results = []
            st.session_state.input_data_list = []
            st.session_state.calculation_history = []
            st.success("✅ Все расчеты очищены!")
        
        if st.button("🧹 Очистить кэш"):
            st.session_state.cache_manager.clear_cache()
            st.success("✅ Кэш очищен!")

def render_help_page():
    """Страница справки"""
    
    st.markdown("## 📖 Справка по FBS юнит-экономике")
    
    st.markdown("""
    ### 🚀 Что такое FBS?
    
    **FBS (Fulfillment by Seller)** — модель работы на маркетплейсах, при которой:
    - Продавец хранит товар на своем складе
    - При поступлении заказа продавец самостоятельно упаковывает и доставляет товар на склад маркетплейса
    - Маркетплейс доставляет товар конечному покупателю (Last Mile)
    
    ### 📊 Ключевые метрики FBS
    
    **1. First Mile (Первая миля)**
    Это ваша логистика от своего склада до склада маркетплейса. Включает:
    - Стоимость транспорта
    - Время на доставку
    - Риски повреждения при транспортировке
    
    **Формула:** `First Mile Cost = (Расстояние × Стоимость 1 км) / Количество единиц на паллете`
    
    **2. Last Mile (Последняя миля)**
    Это логистика маркетплейса от своего склада до клиента. Зависит от:
    - Габаритов товара
    - Веса товара
    - Удаленности клиента
    
    **3. Pick & Pack**
    Стоимость обработки заказа на вашем складе:
    - Время сборки заказа
    - Ставка оператора склада
    - Амортизация оборудования
    
    **Формула:** `Pick & Pack Cost = (Время сборки в минутах / 60) × Ставка оператора в час`
    
    **4. Штрафы за просрочку (Penalty Rate)**
    Маркетплейс устанавливает время, за которое вы должны передать заказ. При просрочке:
    - Штраф составляет определенный процент от цены товара
    - Вероятность просрочки зависит от наличия ночной смены и оперативности
    
    **5. Точка безубыточности по расстоянию**
    Максимальное расстояние до склада МП, при котором First Mile остается рентабельной.
    
    ### 💡 Рекомендации по оптимизации
    
    1. **Минимизируйте First Mile:**
       - Увеличивайте загрузку паллет
       - Оптимизируйте маршруты доставки
       - Рассмотрите FBO для удаленных складов
    
    2. **Снижайте штрафы:**
       - Внедрите ночную смену
       - Автоматизируйте обработку заказов
       - Используйте систему уведомлений
    
    3. **Оптимизируйте Pick & Pack:**
       - Обучайте персонал
       - Внедряйте системы складского учета
       - Оптимизируйте размещение товаров
    
    4. **Работайте с LTV/CAC:**
       - Увеличивайте повторные продажи
       - Оптимизируйте рекламный бюджет
       - Улучшайте качество обслуживания
    """)

# ============================================================================
# БЛОК 9: ГЛАВНАЯ ФУНКЦИЯ ПРИЛОЖЕНИЯ
# ============================================================================

def main():
    """Главная функция приложения"""
    
    # Настройка страницы
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
    
    # Маршрутизация
    current_section = st.session_state.get('current_section', 'main')
    
    if current_section == 'main':
        render_main_page()
    elif current_section == 'calculator':
        render_calculator_page()
    elif current_section == 'batch':
        st.session_state.current_section = 'calculator'
        st.session_state.show_batch = True
        render_calculator_page()
    elif current_section == 'dashboard':
        render_dashboard_page()
    elif current_section == 'export':
        render_export_page()
    elif current_section == 'settings':
        render_settings_page()
    elif current_section == 'help':
        render_help_page()
    else:
        render_main_page()

if __name__ == "__main__":
    main()
