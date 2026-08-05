#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
🚀 ULTIMATE UNIT ECONOMICS FOR YANDEX MARKET v22.1 — REFACTORED & OPTIMIZED
============================================================================
Исправлено:
- Критический баг с st.session_state внутри @st.cache_data (убрана сериализация в JSON)
- Фатальная ошибка в рекомендованных ценах (убрана маскировка знаменателя, добавлен флаг невозможности расчета)
- Побочные эффекты (side effects) в валидаторе данных (теперь работает с копией)
- Оптимизация экспорта Excel: добавлен лимит генерации живых формул (1500 строк) для предотвращения таймаутов
- Восстановлен оборванный текст в UI
"""
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import io
import json
import requests
import logging
import warnings
import hashlib
import re
import time
from enum import Enum
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List, Optional, Tuple, Any
from functools import wraps
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import DataBarRule, FormulaRule
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('YandexMarketUnitEconomics')

APP_VERSION = "22.1.0"
APP_NAME = "Yandex Market Unit Economics PRO"

# ============================================================================
# БЛОК 0: СЛУЖЕБНЫЕ УТИЛИТЫ
# ============================================================================
def money_round(value: float) -> float:
    if pd.isna(value) or np.isinf(value): return 0.0
    return float(Decimal(str(value)).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP))

def percent_round(value: float) -> float:
    if pd.isna(value) or np.isinf(value): return 0.0
    return float(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def safe_divide(numerator: Any, denominator: Any, default: float = 0.0) -> Any:
    """Безопасное деление с защитой от нуля и NaN."""
    denom = np.asarray(denominator, dtype=float)
    num = np.asarray(numerator, dtype=float)
    with np.errstate(divide='ignore', invalid='ignore'):
        result = np.where(np.isclose(denom, 0) | np.isnan(denom) | np.isnan(num), default, num / denom)
    return result

def fix_double_utf8(text: str) -> str:
    if not isinstance(text, str) or not text: return text
    for source_enc, target_enc in [('cp1251', 'utf-8'), ('latin1', 'utf-8')]:
        try:
            fixed = text.encode(source_enc).decode(target_enc)
            if fixed and '\\u0420' not in fixed[:2]: return fixed
        except Exception: continue
    return text

def format_number(num: float, suffix='') -> str:
    if pd.isna(num): return "0"
    abs_num = abs(num)
    sign = "-" if num < 0 else ""
    for unit in ['', 'K', 'M', 'B']:
        if abs_num < 1000.0: return f"{sign}{abs_num:,.1f}{unit}{suffix}".strip()
        abs_num /= 1000.0
    return f"{sign}{abs_num:.1f}T{suffix}"

def make_hash(obj: Any) -> str:
    """Надёжный хеш для pandas DataFrame / dict."""
    try:
        if isinstance(obj, pd.DataFrame):
            return hashlib.sha256(pd.util.hash_pandas_object(obj, index=True).values.tobytes()).hexdigest()[:16]
        return hashlib.sha256(str(obj).encode()).hexdigest()[:16]
    except Exception:
        return hashlib.sha256(str(time.time()).encode()).hexdigest()[:16]

# ============================================================================
# БЛОК 1: КОНФИГУРАЦИИ И МОДЕЛИ
# ============================================================================
class TaxSystem(Enum):
    USN_6 = ("УСН 6% (доходы)", 0.06, "revenue", 0.0)
    USN_15 = ("УСН 15% (доходы-расходы)", 0.15, "profit", 0.01)
    OSN = ("ОСН (общая с НДС 20%)", 0.20, "profit_vat", 0.0)
    AUSN_8 = ("АУСН 8% (доходы)", 0.08, "revenue", 0.0)
    
    def __init__(self, label, rate, base, min_rate):
        self.label = label
        self.rate = rate
        self.base = base
        self.min_rate = min_rate
    
    @classmethod
    def by_label(cls, label: str):
        for item in cls:
            if item.label == label:
                return item
        return cls.USN_6

class YMScheme(Enum):
    FBS = "FBS (склад продавца)"
    FBY = "FBY (склад Маркета)"
    EXPRESS = "Экспресс"
    DBS = "DBS (доставка продавца)"

class Tariff:
    """Модель тарифа с валидацией."""
    def __init__(self, category: str, commission_rate: float = 0.15, min_commission: float = 0.0,
                 sorting_cost: float = 45.0, delivery_rate: float = 0.045, delivery_min: float = 60.0, 
                 delivery_max: float = 500.0, acquiring_transfer_rate: float = 0.016, acquiring_sku_cost: float = 0.12,
                 return_rate: float = 0.05, return_processing: float = 15.0, storage_fee_per_day: float = 0.50,
                 special_tariff_rate: float = 0.42, source: str = "Базовый фоллбэк", scheme: str = "FBS"):
        self.category = str(category).lower().strip()
        self.commission_rate = max(0.0, float(commission_rate))
        self.min_commission = max(0.0, float(min_commission))
        self.sorting_cost = max(0.0, float(sorting_cost))
        self.delivery_rate = max(0.0, float(delivery_rate))
        self.delivery_min = max(0.0, float(delivery_min))
        self.delivery_max = max(float(delivery_min), float(delivery_max))
        self.acquiring_transfer_rate = max(0.0, float(acquiring_transfer_rate))
        self.acquiring_sku_cost = max(0.0, float(acquiring_sku_cost))
        self.return_rate = max(0.0, float(return_rate))
        self.return_processing = max(0.0, float(return_processing))
        self.storage_fee_per_day = max(0.0, float(storage_fee_per_day))
        self.special_tariff_rate = max(0.0, float(special_tariff_rate))
        self.source = source
        self.scheme = scheme

    def to_dict(self) -> Dict[str, Any]:
        return {
            'category': self.category, 'commission_rate': self.commission_rate,
            'min_commission': self.min_commission, 'sorting_cost': self.sorting_cost,
            'delivery_rate': self.delivery_rate, 'delivery_min': self.delivery_min,
            'delivery_max': self.delivery_max, 'acquiring_transfer_rate': self.acquiring_transfer_rate,
            'acquiring_sku_cost': self.acquiring_sku_cost, 'return_rate': self.return_rate,
            'return_processing': self.return_processing, 'storage_fee_per_day': self.storage_fee_per_day,
            'special_tariff_rate': self.special_tariff_rate, 'source': self.source, 'scheme': self.scheme
        }

# ============================================================================
# БЛОК 2: API КЛИЕНТЫ С RETRY / BACKOFF
# ============================================================================
class APIClient:
    """Базовый HTTP-клиент с retry и exponential backoff."""
    def __init__(self, base_url: str, api_key: str, max_retries: int = 3, timeout: int = 15):
        self.base_url = base_url.rstrip('/')
        self.api_key = api_key
        self.timeout = timeout
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
    
    def _request(self, method: str, endpoint: str, **kwargs) -> Dict:
        if not self.api_key:
            return {}
        url = f"{self.base_url}/{endpoint.lstrip('/')}"
        try:
            resp = self.session.request(method, url, timeout=self.timeout, **kwargs)
            resp.raise_for_status()
            return resp.json()
        except requests.exceptions.Timeout:
            logger.warning(f"Таймаут {self.timeout}s при запросе {url}")
        except requests.exceptions.HTTPError as e:
            logger.warning(f"HTTP ошибка {e.response.status_code} при {url}: {e.response.text[:200]}")
        except Exception as e:
            logger.warning(f"Ошибка запроса {url}: {e}")
        return {}

class YandexMarketAPI(APIClient):
    def __init__(self, api_key: str, business_id: Optional[str] = None):
        super().__init__("https://api.partner.market.yandex.ru", api_key, max_retries=3, timeout=15)
        self.business_id = business_id
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        if business_id:
            self.headers["X-Business-Id"] = business_id
    
    def get_campaigns(self) -> List[Dict]:
        data = self._request("GET", "/v2/campaigns", headers=self.headers)
        return data.get("campaigns", [])
    
    def calculate_tariffs(self, offers: List[Dict], campaign_id: Optional[int] = None, 
                          selling_program: str = "FBS") -> List[Dict]:
        payload = {
            "parameters": {
                "sellingProgram": selling_program,
                "frequency": "WEEKLY",
                "paymentDelayWeeks": 4,
                "currency": "RUR"
            },
            "offers": offers
        }
        if campaign_id:
            payload["parameters"]["campaignId"] = campaign_id
            del payload["parameters"]["sellingProgram"]
        data = self._request("POST", "/v2/tariffs/calculate", headers=self.headers, json=payload)
        return data.get("result", {}).get("offers", [])

# ============================================================================
# БЛОК 3: МЕНЕДЖЕР ТАРИФОВ (гибридный)
# ============================================================================
class HybridTariffManager:
    """Управление тарифами: API → Файл → Базовый фоллбэк."""
    
    def __init__(self):
        if 'tariffs' not in st.session_state:
            st.session_state.tariffs = {}
        if 'ym_api_cache' not in st.session_state:
            st.session_state.ym_api_cache = {}
    
    @property
    def tariffs(self) -> Dict[str, Tariff]:
        return st.session_state.tariffs
    
    def load_tariffs_from_file(self, df: pd.DataFrame):
        req_cols = ['category', 'commission_rate']
        if not all(col in df.columns for col in req_cols):
            raise ValueError("Файл тарифов должен содержать минимум: category, commission_rate")
        loaded = 0
        for _, row in df.iterrows():
            cat = str(row['category']).lower().strip()
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
                special_tariff_rate=float(row.get('special_tariff_rate', 0.42)),
                source="Загружено пользователем",
                scheme=str(row.get('scheme', 'FBS'))
            )
            loaded += 1
        logger.info(f"Загружено {loaded} тарифов из файла")

    def get_best_tariff(self, category_name: str, scheme: str,
                        ym_api: Optional[YandexMarketAPI] = None,
                        use_api: bool = True) -> Tariff:
        cat_clean = str(category_name).lower().strip()
        cache_key = f"{cat_clean}_{scheme}"
        
        if cache_key in st.session_state.ym_api_cache:
            return st.session_state.ym_api_cache[cache_key]
        
        # Приоритет 1: API Яндекс Маркета
        if use_api and ym_api and ym_api.api_key:
            try:
                result = ym_api.calculate_tariffs(
                    [{"categoryId": 0, "price": 1000, "length": 10, "width": 10, 
                      "height": 10, "weight": 1, "quantity": 1}],
                    selling_program=scheme
                )
                if result and len(result) > 0:
                    t = self._parse_ym_tariffs(result[0].get("tariffs", []), cat_clean, scheme)
                    if t:
                        st.session_state.ym_api_cache[cache_key] = t
                        return t
            except Exception as e:
                logger.warning(f"API ЯМ сбой для {cat_clean}: {e}")
        
        # Приоритет 2: Загруженный справочник
        if cat_clean in self.tariffs:
            t = self.tariffs[cat_clean]
            t.scheme = scheme
            st.session_state.ym_api_cache[cache_key] = t
            return t
        
        # Фоллбэк
        logger.warning(f"Тариф для '{cat_clean}' не найден. Применён базовый фоллбэк 15%.")
        t = Tariff(category=cat_clean, commission_rate=0.15,
                   source="⚠️ БАЗОВЫЙ ФОЛЛБЭК (требует проверки)", scheme=scheme)
        st.session_state.ym_api_cache[cache_key] = t
        return t
    
    def get_tariffs_vectorized(self, df: pd.DataFrame, scheme: str,
                               ym_api: Optional[YandexMarketAPI] = None,
                               use_api: bool = True) -> pd.DataFrame:
        """Векторизованная подгрузка тарифов для DataFrame."""
        unique_cats = df['category'].dropna().unique()
        tariff_map = {}
        for cat in unique_cats:
            tariff_map[cat] = self.get_best_tariff(cat, scheme, ym_api, use_api)
        
        tariff_df = pd.DataFrame([
            {'category': cat, **t.to_dict()} for cat, t in tariff_map.items()
        ])
        return tariff_df

    def _parse_ym_tariffs(self, tariffs_data: List[Dict], category: str, scheme: str) -> Optional[Tariff]:
        if not tariffs_data:
            return None
        comm_rate, sort_cost, del_rate, acq_rate = 0.15, 45.0, 0.045, 0.016
        for t in tariffs_data:
            t_type, amount = t.get("type", ""), t.get("amount", 0)
            params = {p.get("name", "").lower(): p.get("value", "") for p in t.get("parameters", [])}
            if t_type == "FEE" and params.get("valuetype") == "relative":
                comm_rate = amount / 100.0
            elif t_type == "SORTING":
                sort_cost = amount
            elif t_type == "DELIVERY_TO_CUSTOMER" and params.get("valuetype") == "relative":
                del_rate = amount / 100.0
            elif t_type == "PAYMENT_TRANSFER" and params.get("valuetype") == "relative":
                acq_rate = amount / 100.0
        return Tariff(
            category=category, commission_rate=comm_rate, sorting_cost=sort_cost,
            delivery_rate=del_rate, acquiring_transfer_rate=acq_rate,
            source=f"API Яндекс Маркета ({scheme})", scheme=scheme
        )
    
    def to_dataframe(self) -> pd.DataFrame:
        if not self.tariffs:
            return pd.DataFrame(columns=[
                'Категория', 'Комиссия, %', 'Источник данных', 'Мин. комиссия, ₽',
                'Сортировка, ₽', 'Доставка %', 'Доставка мин, ₽', 'Доставка макс, ₽',
                'Эквайринг перевод, %', 'Эквайринг SKU, ₽', 'Возвраты, %',
                'Обработка возврата, ₽', 'Хранение день, ₽', 'Спецтариф <=300₽, %',
                'Схема'
            ])
        rows = []
        for k, t in self.tariffs.items():
            rows.append({
                'Категория': k,
                'Комиссия, %': round(t.commission_rate * 100, 2),
                'Мин. комиссия, ₽': t.min_commission,
                'Сортировка, ₽': t.sorting_cost,
                'Доставка %': round(t.delivery_rate * 100, 2),
                'Доставка мин, ₽': t.delivery_min,
                'Доставка макс, ₽': t.delivery_max,
                'Эквайринг перевод, %': round(t.acquiring_transfer_rate * 100, 2),
                'Эквайринг SKU, ₽': t.acquiring_sku_cost,
                'Возвраты, %': round(t.return_rate * 100, 2),
                'Обработка возврата, ₽': t.return_processing,
                'Хранение день, ₽': t.storage_fee_per_day,
                'Спецтариф <=300₽, %': round(t.special_tariff_rate * 100, 2),
                'Схема': t.scheme,
                'Источник данных': t.source
            })
        return pd.DataFrame(rows)

# ============================================================================
# БЛОК 4: ВАЛИДАТОР ДАННЫХ
# ============================================================================
class DataValidator:
    REQUIRED_COLS = ['artikul', 'category', 'selling_price', 'cogs']
    NUMERIC_COLS = ['selling_price', 'cogs', 'weight_kg', 'length_cm', 'width_cm', 
                    'height_cm', 'volume_liters', 'packaging_cost', 'first_mile_cost',
                    'marketing_budget_per_unit', 'stock_depth_days', 'quantity_per_order', 'daily_sales']
    
    @classmethod
    def validate(cls, df: pd.DataFrame) -> Tuple[pd.DataFrame, List[str]]:
        errors = []
        if df.empty:
            errors.append("DataFrame пустой")
            return df, errors
        
        # ИСПРАВЛЕНИЕ: Работаем с копией, чтобы избежать побочных эффектов (side effects)
        df_validated = df.copy()
        
        missing = [c for c in cls.REQUIRED_COLS if c not in df_validated.columns]
        if missing:
            errors.append(f"Отсутствуют обязательные колонки: {missing}")
        
        for col in cls.NUMERIC_COLS:
            if col in df_validated.columns:
                negatives = (df_validated[col] < 0).sum()
                if negatives > 0:
                    errors.append(f"{col}: {negatives} отрицательных значений исправлено на 0")
                    df_validated[col] = df_validated[col].clip(lower=0)
        
        if 'selling_price' in df_validated.columns:
            zero_prices = (df_validated['selling_price'] == 0).sum()
            if zero_prices > 0:
                errors.append(f"selling_price: {zero_prices} SKU с нулевой ценой (расчёт будет некорректным)")
        
        if 'quantity_per_order' in df_validated.columns:
            zero_qty = (df_validated['quantity_per_order'] == 0).sum()
            if zero_qty > 0:
                errors.append(f"quantity_per_order: {zero_qty} нулевых значений исправлено на 1")
                df_validated['quantity_per_order'] = df_validated['quantity_per_order'].replace(0, 1)
        
        return df_validated, errors

# ============================================================================
# БЛОК 5: ВЕКТОРИЗОВАННЫЙ ФИНАНСОВЫЙ ДВИЖОК
# ============================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def run_calculations_cached(
    df: pd.DataFrame,
    tax_label: str,
    scheme_label: str,
    payment_frequency: str,
    tariffs_map: dict
) -> pd.DataFrame:
    """
    Кешированный расчёт unit-экономики.
    ИСПРАВЛЕНИЕ: Принимаем нативный pd.DataFrame и dict, без сериализации в JSON.
    """
    if df.empty:
        return df
    
    tax_system = TaxSystem.by_label(tax_label)
    scheme = scheme_label.split(" ")[0]
    
    payment_rates = {
        "Ежемесячно (1.0%)": 0.01,
        "Раз в 2 недели (1.3%)": 0.013,
        "Еженедельно, 4 нед. (1.6%)": 0.016,
        "Ежедневно (3.3%)": 0.033
    }
    p_transfer_rate = payment_rates.get(payment_frequency, 0.016)
    
    # Обработка текстовых полей
    for col in ['artikul', 'category']:
        if col in df.columns:
            df[col] = df[col].astype(str).apply(fix_double_utf8)
    
    # Заполнение умолчаний
    defaults = {
        'selling_price': 0.0, 'cogs': 0.0, 'weight_kg': 0.0, 'length_cm': 0.0,
        'width_cm': 0.0, 'height_cm': 0.0, 'packaging_cost': 0.0,
        'marketing_budget_per_unit': 0.0, 'daily_sales': 0.0, 'stock_depth_days': 0.0,
        'first_mile_cost': 0.0, 'commission': 0.0, 'return_cost': 0.0,
        'warehouse_cost': 0.0, 'volume_liters': 0.0, 'quantity_per_order': 1.0
    }
    for col, default in defaults.items():
        if col not in df.columns:
            df[col] = default
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(default)
    
    # === ВЕКТОРИЗОВАННАЯ ПОДГРУЗКА ТАРИФОВ ===
    tariff_df = pd.DataFrame.from_dict(tariffs_map, orient='index')
    if tariff_df.empty:
        tariff_df = pd.DataFrame([Tariff('default').to_dict()])
    
    # Мержим тарифы по категории
    df = df.merge(tariff_df, on='category', how='left')
    
    # Фоллбэк для неизвестных категорий
    for col in ['commission_rate', 'sorting_cost', 'delivery_rate', 'delivery_min',
                'delivery_max', 'acquiring_transfer_rate', 'acquiring_sku_cost',
                'return_rate', 'return_processing', 'storage_fee_per_day', 'special_tariff_rate']:
        if col in df.columns:
            df[col] = df[col].fillna(Tariff('default').to_dict()[col])
    
    # === РАСЧЁТЫ ===
    vol_weight = (df['length_cm'] * df['width_cm'] * df['height_cm']) / 5000.0
    df['billable_weight'] = np.ceil(np.maximum(df['weight_kg'], vol_weight) * 2) / 2
    df['is_special_tariff'] = (df['selling_price'] <= 300) & (df['volume_liters'] <= 5)
    
    # Комиссия
    df['commission'] = np.where(
        df['is_special_tariff'],
        df['selling_price'] * df['special_tariff_rate'],
        np.maximum(df['selling_price'] * df['commission_rate'], df['min_commission'])
    )
    
    # Доставка покупателю (с порогами)
    raw_delivery = df['selling_price'] * df['delivery_rate']
    df['delivery_to_customer'] = np.where(
        df['is_special_tariff'],
        0.0,
        np.clip(raw_delivery, df['delivery_min'], df['delivery_max'])
    )
    
    # Средняя миля
    df['middle_mile_cost'] = np.where(
        df['is_special_tariff'],
        0.0,
        np.where(df['billable_weight'] <= 4, 100,
                 np.where(df['billable_weight'] <= 10, 300, 600))
    )
    
    # Сортировка
    df['sorting_cost'] = np.where(
        df['is_special_tariff'],
        0.0,
        np.where(scheme == 'FBS', df['sorting_cost'], 0.0)
    )
    
    # Эквайринг
    df['acquiring_sku_cost'] = df['acquiring_sku_cost'] / df['quantity_per_order']
    df['acquiring_transfer_cost'] = df['selling_price'] * p_transfer_rate
    df['acquiring_cost'] = df['acquiring_sku_cost'] + df['acquiring_transfer_cost']
    
    # Возвраты
    df['return_processing_cost'] = np.where(df['is_special_tariff'], 0.0, df['return_processing'])
    df['return_delivery_cost'] = np.where(df['is_special_tariff'], 0.0, df['middle_mile_cost'] * df['return_rate'])
    df['return_cost'] = df['return_processing_cost'] + df['return_delivery_cost']
    
    # Упаковка + склад
    df['pick_pack_cost'] = 35.0
    df['warehouse_cost'] = np.where(
        df['warehouse_cost'] == 0,
        (df['stock_depth_days'] * df['daily_sales']) * df['storage_fee_per_day'],
        df['warehouse_cost']
    )
    
    # Итоговые расходы
    df['fixed_operational_costs'] = (
        df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] +
        df['packaging_cost'] + df['return_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    )
    df['marketplace_fees'] = (
        df['commission'] + df['delivery_to_customer'] + df['middle_mile_cost'] +
        df['sorting_cost'] + df['acquiring_cost']
    )
    df['pre_tax_expenses'] = df['fixed_operational_costs'] + df['marketplace_fees']
    
    # Налог
    if tax_system.base == "revenue":
        df['tax_cost'] = df['selling_price'] * tax_system.rate
    elif tax_system.base == "profit_vat":
        revenue_without_vat = df['selling_price'] / 1.20
        pre_tax_profit = revenue_without_vat - df['pre_tax_expenses']
        df['tax_cost'] = np.maximum(pre_tax_profit, 0) * tax_system.rate
    else:  # profit
        pre_tax_profit = df['selling_price'] - df['pre_tax_expenses']
        df['tax_cost'] = np.maximum(pre_tax_profit, 0) * tax_system.rate
    
    df['total_expenses'] = df['pre_tax_expenses'] + df['tax_cost']
    df['gross_profit'] = df['selling_price'] - df['total_expenses']
    df['margin_percent'] = np.where(df['selling_price'] > 0, (df['gross_profit'] / df['selling_price']) * 100, 0.0)
    
    # Рекомендованные цены (ИСПРАВЛЕНИЕ: честная обработка фатальной убыточности)
    var_fees = np.where(
        df['is_special_tariff'],
        df['special_tariff_rate'] + p_transfer_rate + (tax_system.rate if tax_system.base == "revenue" else 0),
        df['commission_rate'] + df['delivery_rate'] + p_transfer_rate + (tax_system.rate if tax_system.base == "revenue" else 0)
    )
    denom = 1.0 - var_fees
    denom_valid = denom > 0.01  # Если издержки >= 99%, модель фатально убыточна
    
    fixed_no_return = (
        df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] +
        df['packaging_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    )
    
    df['rec_price_min'] = np.where(denom_valid, safe_divide(fixed_no_return, denom, default=np.nan), np.nan)
    df['rec_price_15'] = np.where(denom_valid, safe_divide(fixed_no_return, denom - 0.15, default=np.nan), np.nan)
    df['rec_price_25'] = np.where(denom_valid, safe_divide(fixed_no_return, denom - 0.25, default=np.nan), np.nan)
    df['is_price_calc_possible'] = denom_valid
    
    # Доп. метрики
    df['variable_costs'] = (
        df['commission'] + df['delivery_to_customer'] + df['middle_mile_cost'] +
        df['sorting_cost'] + df['acquiring_cost'] + df['return_cost']
    )
    df['fixed_costs'] = fixed_no_return + df['return_cost']
    df['contribution_margin'] = df['selling_price'] - df['variable_costs']
    df['roi_percent'] = np.where(df['cogs'] > 0, (df['gross_profit'] / df['cogs']) * 100, 0.0)
    df['break_even_units'] = safe_divide(df['fixed_costs'], df['contribution_margin'], default=0.0)
    
    # ABC-XYZ
    df['abc_category'] = np.where(df['daily_sales'] >= 10, 'A', np.where(df['daily_sales'] >= 3, 'B', 'C'))
    df['xyz_category'] = np.where(df['margin_percent'] >= 20, 'X', np.where(df['margin_percent'] >= 10, 'Y', 'Z'))
    df['abc_xyz'] = df['abc_category'] + df['xyz_category']
    df['profitability_status'] = np.where(
        df['gross_profit'] > 0,
        np.where(df['margin_percent'] >= 20, 'Высокомаржинальный', 'Низкомаржинальный'),
        'Убыточный'
    )
    
    # Округление
    money_cols = ['commission', 'delivery_to_customer', 'middle_mile_cost', 'sorting_cost',
                  'acquiring_cost', 'return_cost', 'gross_profit', 'total_expenses',
                  'rec_price_min', 'rec_price_15', 'rec_price_25', 'tax_cost']
    for col in money_cols:
        if col in df.columns:
            df[col] = df[col].apply(money_round)
    
    pct_cols = ['margin_percent', 'roi_percent']
    for col in pct_cols:
        if col in df.columns:
            df[col] = df[col].apply(percent_round)
    
    # Убираем технические колонки тарифов
    tech_cols = ['commission_rate', 'min_commission', 'sorting_cost', 'delivery_rate',
                 'delivery_min', 'delivery_max', 'acquiring_transfer_rate', 'acquiring_sku_cost',
                 'return_rate', 'return_processing', 'storage_fee_per_day', 'special_tariff_rate',
                 'source', 'scheme']
    for col in tech_cols:
        if col in df.columns:
            df = df.drop(columns=[col])
    
    return df

# ============================================================================
# БЛОК 6: ЭКСПОРТЁР EXCEL С ЖИВЫМИ ФОРМУЛАМИ (ОПТИМИЗИРОВАННЫЙ)
# ============================================================================
class UltimateExcelExporter:
    """Экспорт с корректными VLOOKUP, MIN/MAX, IF и условным форматированием."""
    
    TARIFF_COLS = [
        'Категория', 'Комиссия, %', 'Мин. комиссия, ₽', 'Сортировка, ₽',
        'Доставка %', 'Доставка мин, ₽', 'Доставка макс, ₽',
        'Эквайринг перевод, %', 'Эквайринг SKU, ₽', 'Возвраты, %',
        'Обработка возврата, ₽', 'Хранение день, ₽', 'Спецтариф <=300₽, %',
        'Схема', 'Источник данных'
    ]
    
    # Индексы VLOOKUP (1-based)
    VLOOKUP_IDX = {
        'commission_rate': 2,      # B
        'min_commission': 3,       # C
        'sorting_cost': 4,         # D
        'delivery_rate': 5,        # E
        'delivery_min': 6,         # F
        'delivery_max': 7,         # G
        'acquiring_transfer': 8,   # H
        'acquiring_sku': 9,        # I
        'return_rate': 10,         # J
        'return_processing': 11,   # K
        'storage_fee': 12,         # L
        'special_tariff': 13,      # M
    }

    @staticmethod
    def export_max_info(df: pd.DataFrame, tax_label: str, scheme_label: str,
                        tariff_manager: HybridTariffManager,
                        payment_frequency: str) -> bytes:
        if not OPENPYXL_AVAILABLE or df.empty:
            return b""
        
        wb = Workbook()
        
        # === 0. ЛИСТ ПАРАМЕТРОВ (глобальные настройки для формул) ===
        ws_params = wb.active
        ws_params.title = "Параметры"
        tax_system = TaxSystem.by_label(tax_label)
        scheme = scheme_label.split(" ")[0]
        payment_rates = {
            "Ежемесячно (1.0%)": 1.0, "Раз в 2 недели (1.3%)": 1.3,
            "Еженедельно, 4 нед. (1.6%)": 1.6, "Ежедневно (3.3%)": 3.3
        }
        acq_transfer_pct = payment_rates.get(payment_frequency, 1.6)
        
        params_data = [
            ("Налоговая система", tax_label),
            ("Ставка налога", tax_system.rate),
            ("База налога", tax_system.base),
            ("Схема", scheme),
            ("Эквайринг перевод %", acq_transfer_pct / 100),
        ]
        for r, (label, val) in enumerate(params_data, 1):
            ws_params.cell(r, 1, label).font = Font(bold=True)
            c = ws_params.cell(r, 2, val)
            if isinstance(val, float):
                c.number_format = '0.00%' if 'ставка' in label.lower() else '0.00'
        ws_params.column_dimensions['A'].width = 25
        ws_params.column_dimensions['B'].width = 30
        
        # === 1. ЛИСТ СПРАВОЧНИКА ТАРИФОВ ===
        ws_tariffs = wb.create_sheet("Тарифы_Справочник")
        tariff_df = tariff_manager.to_dataframe()
        if tariff_df.empty:
            tariff_df = pd.DataFrame([{
                'Категория': 'default', 'Комиссия, %': 15.0, 'Мин. комиссия, ₽': 0,
                'Сортировка, ₽': 45, 'Доставка %': 4.5, 'Доставка мин, ₽': 60,
                'Доставка макс, ₽': 500, 'Эквайринг перевод, %': 1.6,
                'Эквайринг SKU, ₽': 0.12, 'Возвраты, %': 5.0,
                'Обработка возврата, ₽': 15, 'Хранение день, ₽': 0.5,
                'Спецтариф <=300₽, %': 42.0, 'Схема': 'FBS',
                'Источник данных': 'Базовый фоллбэк'
            }])
        
        for c_idx, c_name in enumerate(UltimateExcelExporter.TARIFF_COLS, 1):
            cell = ws_tariffs.cell(1, c_idx, c_name)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="1F4E78", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        
        for r_idx, row_data in enumerate(tariff_df.itertuples(index=False), 2):
            for c_idx, value in enumerate(row_data, 1):
                ws_tariffs.cell(r_idx, c_idx, value)
        
        for c_idx in range(1, len(UltimateExcelExporter.TARIFF_COLS) + 1):
            ws_tariffs.column_dimensions[get_column_letter(c_idx)].width = 18
        
        # === 2. ДЕТАЛЬНЫЙ РАСЧЁТ С ЖИВЫМИ ФОРМУЛАМИ ===
        ws_detail = wb.create_sheet("Детальный_Расчет")
        
        priority_cols = [
            'artikul', 'category', 'selling_price', 'cogs', 'weight_kg', 'length_cm',
            'width_cm', 'height_cm', 'volume_liters', 'quantity_per_order',
            'billable_weight', 'is_special_tariff', 'commission', 'delivery_to_customer',
            'middle_mile_cost', 'sorting_cost', 'acquiring_sku_cost', 'acquiring_transfer_cost',
            'acquiring_cost', 'return_processing_cost', 'return_delivery_cost', 'return_cost',
            'pick_pack_cost', 'warehouse_cost', 'marketing_budget_per_unit', 'first_mile_cost',
            'packaging_cost', 'fixed_operational_costs', 'marketplace_fees',
            'pre_tax_expenses', 'tax_cost', 'total_expenses', 'gross_profit',
            'margin_percent', 'rec_price_min', 'rec_price_15', 'rec_price_25',
            'profitability_status', 'abc_xyz', 'daily_sales', 'is_price_calc_possible'
        ]
        cols = [c for c in priority_cols if c in df.columns] + [c for c in df.columns if c not in priority_cols]
        
        # Заголовки
        for c_idx, c_name in enumerate(cols, 1):
            cell = ws_detail.cell(1, c_idx, c_name)
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = PatternFill(start_color="2E75B6", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = Border(
                left=Side(style="thin"), right=Side(style="thin"),
                top=Side(style="thin"), bottom=Side(style="thin")
            )
        
        ws_detail.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{len(df)+1}"
        ws_detail.freeze_panes = 'B2'
        
        # ИСПРАВЛЕНИЕ: Оптимизация поиска индексов (выносим из цикла)
        idx_map = {col: i for i, col in enumerate(cols)}
        col_letter_map = {col: get_column_letter(i + 1) for col, i in idx_map.items()}
        
        def vlookup(col_name: str, lookup_cell: str) -> str:
            idx = UltimateExcelExporter.VLOOKUP_IDX.get(col_name, 2)
            return f"IFERROR(VLOOKUP({lookup_cell}, 'Тарифы_Справочник'!$A:$O, {idx}, FALSE), 0)"
        
        # ИСПРАВЛЕНИЕ: Лимит генерации формул для предотвращения фатальных таймаутов Streamlit
        MAX_FORMULA_ROWS = 1500
        use_live_formulas = len(df) <= MAX_FORMULA_ROWS
        
        if not use_live_formulas:
            warning_cell = ws_detail.cell(1, len(cols) + 1, "⚠️ Живые формулы отключены для датасетов > 1500 строк во избежание сбоев. Используются рассчитанные значения.")
            warning_cell.font = Font(color="FF0000", bold=True)
        
        for r_idx in range(2, len(df) + 2):
            # Записываем входные данные и значения (быстро)
            for c_idx, col_name in enumerate(cols, 1):
                if col_name in df.columns:
                    val = df.iloc[r_idx - 2][col_name]
                    cell = ws_detail.cell(r_idx, c_idx, val)
                    cell.border = Border(bottom=Side(style="thin", color="E0E0E0"))
                    if isinstance(val, (int, float)) and not pd.isna(val):
                        if 'percent' in col_name:
                            cell.number_format = '0.00"%"'
                        elif col_name in ['daily_sales', 'break_even_units', 'quantity_per_order']:
                            cell.number_format = '0'
                        else:
                            cell.number_format = '#,##0.00'
            
            # Генерируем живые формулы только если датасет в безопасных пределах
            if use_live_formulas:
                cat = f"{col_letter_map.get('category', 'B')}{r_idx}"
                price = f"{col_letter_map.get('selling_price', 'C')}{r_idx}"
                spec = f"{col_letter_map.get('is_special_tariff', 'L')}{r_idx}"
                bw = f"{col_letter_map.get('billable_weight', 'K')}{r_idx}"
                
                if 'billable_weight' in idx_map:
                    w = col_letter_map.get('weight_kg', 'E')
                    l = col_letter_map.get('length_cm', 'F')
                    wdt = col_letter_map.get('width_cm', 'G')
                    h = col_letter_map.get('height_cm', 'H')
                    ws_detail.cell(r_idx, idx_map['billable_weight'] + 1).value = (
                        f"=ROUNDUP(MAX({w}{r_idx}, {l}{r_idx}*{wdt}{r_idx}*{h}{r_idx}/5000)*2,0)/2"
                    )
                
                if 'is_special_tariff' in idx_map:
                    v = col_letter_map.get('volume_liters', 'I')
                    ws_detail.cell(r_idx, idx_map['is_special_tariff'] + 1).value = (
                        f"=IF(AND({price}<=300, {v}{r_idx}<=5), 1, 0)"
                    )
                
                if 'commission' in idx_map:
                    ws_detail.cell(r_idx, idx_map['commission'] + 1).value = (
                        f"=IF({spec}=1, {price}*{vlookup('special_tariff', cat)}/100, "
                        f"MAX({price}*{vlookup('commission_rate', cat)}/100, {vlookup('min_commission', cat)}))"
                    )
                
                if 'delivery_to_customer' in idx_map:
                    ws_detail.cell(r_idx, idx_map['delivery_to_customer'] + 1).value = (
                        f"=IF({spec}=1, 0, MIN(MAX({price}*{vlookup('delivery_rate', cat)}/100, "
                        f"{vlookup('delivery_min', cat)}), {vlookup('delivery_max', cat)}))"
                    )
                
                if 'middle_mile_cost' in idx_map:
                    ws_detail.cell(r_idx, idx_map['middle_mile_cost'] + 1).value = (
                        f"=IF({spec}=1, 0, IF({bw}<=4, 100, IF({bw}<=10, 300, 600)))"
                    )
                
                if 'sorting_cost' in idx_map:
                    ws_detail.cell(r_idx, idx_map['sorting_cost'] + 1).value = (
                        f"=IF({spec}=1, 0, IF(Параметры!$B$4=\"FBS\", {vlookup('sorting_cost', cat)}, 0))"
                    )
                
                if 'acquiring_sku_cost' in idx_map:
                    qty = col_letter_map.get('quantity_per_order', 'J')
                    ws_detail.cell(r_idx, idx_map['acquiring_sku_cost'] + 1).value = (
                        f"={vlookup('acquiring_sku', cat)}/{qty}{r_idx}"
                    )
                
                if 'acquiring_transfer_cost' in idx_map:
                    ws_detail.cell(r_idx, idx_map['acquiring_transfer_cost'] + 1).value = (
                        f"={price}*Параметры!$B$5"
                    )
                
                if 'acquiring_cost' in idx_map:
                    ask = col_letter_map.get('acquiring_sku_cost', 'O')
                    atr = col_letter_map.get('acquiring_transfer_cost', 'P')
                    ws_detail.cell(r_idx, idx_map['acquiring_cost'] + 1).value = (
                        f"={ask}{r_idx}+{atr}{r_idx}"
                    )
                
                if 'return_processing_cost' in idx_map:
                    ws_detail.cell(r_idx, idx_map['return_processing_cost'] + 1).value = (
                        f"=IF({spec}=1, 0, {vlookup('return_processing', cat)})"
                    )
                
                if 'return_delivery_cost' in idx_map:
                    mm = col_letter_map.get('middle_mile_cost', 'N')
                    ws_detail.cell(r_idx, idx_map['return_delivery_cost'] + 1).value = (
                        f"=IF({spec}=1, 0, {mm}{r_idx}*{vlookup('return_rate', cat)}/100)"
                    )
                
                if 'return_cost' in idx_map:
                    rpc = col_letter_map.get('return_processing_cost', 'T')
                    rdc = col_letter_map.get('return_delivery_cost', 'U')
                    ws_detail.cell(r_idx, idx_map['return_cost'] + 1).value = (
                        f"={rpc}{r_idx}+{rdc}{r_idx}"
                    )
                
                if 'pick_pack_cost' in idx_map:
                    ws_detail.cell(r_idx, idx_map['pick_pack_cost'] + 1).value = 35.0
                
                if 'fixed_operational_costs' in idx_map:
                    cogs = col_letter_map.get('cogs', 'D')
                    fmc = col_letter_map.get('first_mile_cost', 'M')
                    ppc = col_letter_map.get('pick_pack_cost', 'V')
                    pck = col_letter_map.get('packaging_cost', 'W')
                    rc = col_letter_map.get('return_cost', 'X')
                    mkt = col_letter_map.get('marketing_budget_per_unit', 'Y')
                    wh = col_letter_map.get('warehouse_cost', 'Z')
                    ws_detail.cell(r_idx, idx_map['fixed_operational_costs'] + 1).value = (
                        f"={cogs}{r_idx}+{fmc}{r_idx}+{ppc}{r_idx}+{pck}{r_idx}+{rc}{r_idx}+{mkt}{r_idx}+{wh}{r_idx}"
                    )
                
                if 'marketplace_fees' in idx_map:
                    comm = col_letter_map.get('commission', 'M')
                    dlv = col_letter_map.get('delivery_to_customer', 'N')
                    mm = col_letter_map.get('middle_mile_cost', 'O')
                    srt = col_letter_map.get('sorting_cost', 'P')
                    acq = col_letter_map.get('acquiring_cost', 'Q')
                    ws_detail.cell(r_idx, idx_map['marketplace_fees'] + 1).value = (
                        f"={comm}{r_idx}+{dlv}{r_idx}+{mm}{r_idx}+{srt}{r_idx}+{acq}{r_idx}"
                    )
                
                if 'pre_tax_expenses' in idx_map:
                    foc = col_letter_map.get('fixed_operational_costs', 'AA')
                    mf = col_letter_map.get('marketplace_fees', 'AB')
                    ws_detail.cell(r_idx, idx_map['pre_tax_expenses'] + 1).value = (
                        f"={foc}{r_idx}+{mf}{r_idx}"
                    )
                
                if 'tax_cost' in idx_map:
                    pte = col_letter_map.get('pre_tax_expenses', 'AC')
                    ws_detail.cell(r_idx, idx_map['tax_cost'] + 1).value = (
                        f"=IF(Параметры!$B$3=\"revenue\", {price}*Параметры!$B$2, "
                        f"IF(Параметры!$B$3=\"profit_vat\", MAX({price}/1.2-{pte}{r_idx},0)*Параметры!$B$2, "
                        f"MAX({price}-{pte}{r_idx},0)*Параметры!$B$2))"
                    )
                
                if 'total_expenses' in idx_map:
                    pte = col_letter_map.get('pre_tax_expenses', 'AC')
                    tc = col_letter_map.get('tax_cost', 'AD')
                    ws_detail.cell(r_idx, idx_map['total_expenses'] + 1).value = (
                        f"={pte}{r_idx}+{tc}{r_idx}"
                    )
                
                if 'gross_profit' in idx_map:
                    te = col_letter_map.get('total_expenses', 'AE')
                    ws_detail.cell(r_idx, idx_map['gross_profit'] + 1).value = (
                        f"={price}-{te}{r_idx}"
                    )
                
                if 'margin_percent' in idx_map:
                    gp = col_letter_map.get('gross_profit', 'AF')
                    cell = ws_detail.cell(r_idx, idx_map['margin_percent'] + 1)
                    cell.value = f"=IF({price}>0, {gp}{r_idx}/{price}, 0)"
                    cell.number_format = '0.00%'
                
                if 'profitability_status' in idx_map:
                    gp = col_letter_map.get('gross_profit', 'AF')
                    mp = col_letter_map.get('margin_percent', 'AG')
                    ws_detail.cell(r_idx, idx_map['profitability_status'] + 1).value = (
                        f"=IF({gp}{r_idx}>0, IF({mp}{r_idx}>=0.2, \"Высокомаржинальный\", \"Низкомаржинальный\"), \"Убыточный\")"
                    )
        
        # Условное форматирование
        if 'gross_profit' in idx_map:
            col_l = col_letter_map['gross_profit']
            ws_detail.conditional_formatting.add(
                f'{col_l}2:{col_l}{len(df)+1}',
                FormulaRule(formula=[f'{col_l}2>=0'], fill=PatternFill(start_color="C6EFCE"), font=Font(color="006100", bold=True))
            )
            ws_detail.conditional_formatting.add(
                f'{col_l}2:{col_l}{len(df)+1}',
                FormulaRule(formula=[f'{col_l}2<0'], fill=PatternFill(start_color="FFC7CE"), font=Font(color="9C0006", bold=True))
            )
        
        if 'margin_percent' in idx_map:
            col_l = col_letter_map['margin_percent']
            ws_detail.conditional_formatting.add(
                f'{col_l}2:{col_l}{len(df)+1}',
                DataBarRule(start_type='min', end_type='max', color="638EC6", showValue=True)
            )
        
        # Ширины колонок
        for c_idx, c_name in enumerate(cols, 1):
            width = 22 if c_name in ['artikul', 'category', 'profitability_status'] else (14 if 'percent' in c_name else 18)
            ws_detail.column_dimensions[get_column_letter(c_idx)].width = width
        
        # === 3. ДАШБОРД-СВОДКА ===
        ws_dash = wb.create_sheet("Дашборд_Сводка")
        ws_dash.merge_cells('A1:E1')
        cell = ws_dash.cell(1, 1, "СВОДНЫЙ ФИНАНСОВЫЙ ДАШБОРД")
        cell.font = Font(size=16, bold=True, color="1F4E78")
        cell.alignment = Alignment(horizontal="center")
        
        metrics = [
            ("Всего SKU", len(df)),
            ("Общая выручка", df['selling_price'].sum()),
            ("ОБЩАЯ ПРИБЫЛЬ", df['gross_profit'].sum()),
            ("Средняя маржа %", df['margin_percent'].mean() / 100),
        ]
        for r_idx, (label, val) in enumerate(metrics, 3):
            ws_dash.cell(r_idx, 1, label).font = Font(bold=True)
            c = ws_dash.cell(r_idx, 2, val)
            if 'маржа' in label.lower():
                c.number_format = '0.00%'
            else:
                c.number_format = '#,##0.00 "₽"'
            c.font = Font(bold=True, color="1F4E78")
        
        expense_labels = ["Себестоимость", "Комиссия", "Доставка", "Ср. миля", "Эквайринг", "Налоги"]
        expense_vals = [
            df['cogs'].sum(), df['commission'].sum(), df['delivery_to_customer'].sum(),
            df['middle_mile_cost'].sum(), df['acquiring_cost'].sum(), df['tax_cost'].sum()
        ]
        ws_dash.cell(10, 1, "Структура расходов")
        ws_dash.cell(10, 1).font = Font(bold=True, size=12)
        for i, (lbl, val) in enumerate(zip(expense_labels, expense_vals), 11):
            ws_dash.cell(i, 1, lbl)
            ws_dash.cell(i, 2, val)
        
        pie = PieChart()
        pie.title = "Структура расходов"
        labels = Reference(ws_dash, min_col=1, min_row=11, max_row=16)
        data = Reference(ws_dash, min_col=2, min_row=11, max_row=16)
        pie.add_data(data, titles_from_data=False)
        pie.set_categories(labels)
        pie.dataLabels = DataLabelList()
        pie.dataLabels.showPercent = True
        ws_dash.add_chart(pie, "D10")
        
        # === 4. ABC-XYZ АНАЛИЗ ===
        ws_abc = wb.create_sheet("ABC_XYZ")
        abc_data = df.groupby('abc_xyz').agg({
            'artikul': 'count', 'selling_price': 'sum', 'gross_profit': 'sum', 'margin_percent': 'mean'
        }).reset_index()
        abc_data.columns = ['ABC-XYZ', 'Кол-во SKU', 'Выручка', 'Прибыль', 'Ср. маржа %']
        
        for c_idx, c_name in enumerate(abc_data.columns, 1):
            cell = ws_abc.cell(1, c_idx, c_name)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="2E75B6", fill_type="solid")
        
        for r_idx, row_data in enumerate(abc_data.itertuples(index=False), 2):
            for c_idx, val in enumerate(row_data, 1):
                cell = ws_abc.cell(r_idx, c_idx, val)
                if c_idx > 2:
                    cell.number_format = '#,##0.00'
                elif c_idx == 5:
                    cell.number_format = '0.00"%"'
        
        bar = BarChart()
        bar.title = "Прибыль по ABC-XYZ сегментам"
        bar.x_axis.title = "Сегмент"
        bar.y_axis.title = "Прибыль, ₽"
        cats = Reference(ws_abc, min_col=1, min_row=2, max_row=len(abc_data)+1)
        vals = Reference(ws_abc, min_col=4, min_row=1, max_row=len(abc_data)+1)
        bar.add_data(vals, titles_from_data=True)
        bar.set_categories(cats)
        ws_abc.add_chart(bar, "G2")
        
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.getvalue()

# ============================================================================
# БЛОК 7: НОРМАЛИЗАТОР ДАННЫХ
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
        'first_mile_cost': ['first_mile_cost', 'магистраль', 'первая миля', 'доставка до склада'],
        'marketing_budget_per_unit': ['marketing_budget_per_unit', 'реклама', 'дрр', 'маркетинг'],
        'stock_depth_days': ['stock_depth_days', 'глубина запаса', 'дни запаса', 'запас, дни'],
        'quantity_per_order': ['quantity_per_order', 'количество в заказе', 'шт в заказе', 'qty_per_order'],
        'warehouse_cost': ['warehouse_cost', 'стоимость хранения', 'хранение'],
    }
    
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
        
        # Чистка числовых колонок
        num_cols = ['selling_price', 'cogs', 'daily_sales', 'weight_kg', 'length_cm',
                    'width_cm', 'height_cm', 'volume_liters', 'packaging_cost',
                    'first_mile_cost', 'marketing_budget_per_unit', 'stock_depth_days',
                    'quantity_per_order', 'warehouse_cost']
        for col in num_cols:
            if col in norm_df.columns:
                norm_df[col] = pd.to_numeric(
                    norm_df[col].astype(str).str.replace(r'[\s,;%₽]', '', regex=True),
                    errors='coerce'
                ).fillna(0.0).abs()
        
        norm_df['artikul'] = norm_df['artikul'].astype(str).str.strip()
        norm_df['category'] = norm_df['category'].astype(str).str.strip().str.lower()
        
        # Удаляем полные дубликаты артикулов (оставляем первый)
        norm_df = norm_df.drop_duplicates(subset=['artikul'], keep='first')
        
        return norm_df

    @classmethod
    def load_file(cls, file_buffer: io.BytesIO, file_name: str) -> pd.DataFrame:
        try:
            if file_name.endswith('.csv'):
                return pd.read_csv(file_buffer, sep=None, engine='python', encoding='utf-8')
            elif file_name.endswith(('.xls', '.xlsx')):
                return pd.read_excel(file_buffer)
            else:
                raise ValueError("Неподдерживаемый формат. Используйте CSV или XLSX.")
        except UnicodeDecodeError:
            file_buffer.seek(0)
            return pd.read_csv(file_buffer, sep=None, engine='python', encoding='cp1251')

# ============================================================================
# БЛОК 8: STREAMLIT UI
# ============================================================================
def init_session_state():
    defaults = {
        'main_df': pd.DataFrame(),
        'calc_df': pd.DataFrame(),
        'tariffs': {},
        'ym_api_cache': {},
        'scenario_df': pd.DataFrame(),
        'last_hash': '',
    }
    for key, val in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = val

def render_sidebar() -> Tuple[str, str, str, bool, Optional[YandexMarketAPI]]:
    st.sidebar.title("⚙️ Панель управления")
    st.sidebar.markdown(f"**{APP_NAME} v{APP_VERSION}**")
    st.sidebar.markdown("---")
    
    with st.sidebar.form("api_form"):
        st.subheader("🔐 API Доступы")
        api_key = st.text_input(
            "API Key Яндекс Маркета", type="password",
            value=st.secrets.get("MARKET_API_KEY", "")
        )
        business_id = st.text_input(
            "Business ID",
            value=st.secrets.get("MARKET_BUSINESS_ID", "")
        )
        use_api = st.checkbox("🌐 Использовать API для тарифов", value=True)
        st.form_submit_button("💾 Сохранить настройки API")
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("🏪 Настройки бизнеса")
    scheme_label = st.sidebar.selectbox("Схема работы:", [s.value for s in YMScheme])
    tax_label = st.sidebar.selectbox("Налогообложение:", [t.label for t in TaxSystem])
    payment_freq = st.sidebar.selectbox(
        "Частота выплат:",
        ["Ежемесячно (1.0%)", "Раз в 2 недели (1.3%)",
         "Еженедельно, 4 нед. (1.6%)", "Ежедневно (3.3%)"],
        index=2
    )
    
    ym_api = None
    if api_key and use_api:
        ym_api = YandexMarketAPI(api_key, business_id if business_id else None)
    
    return scheme_label, tax_label, payment_freq, use_api, ym_api

def run_analysis(df: pd.DataFrame, tax_label: str, scheme_label: str,
                 payment_freq: str, tm: HybridTariffManager,
                 ym_api: Optional[YandexMarketAPI], use_api: bool) -> pd.DataFrame:
    """Запуск расчёта с прогресс-баром и валидацией."""
    if df.empty:
        return pd.DataFrame()
    
    with st.spinner("🔍 Валидация данных..."):
        df_validated, errors = DataValidator.validate(df)
        if errors:
            for err in errors:
                st.warning(err)
    
    with st.spinner("📡 Подгрузка тарифов..."):
        tariff_df = tm.get_tariffs_vectorized(df_validated, scheme_label.split(" ")[0], ym_api, use_api)
        df_merged = df_validated.merge(tariff_df, on='category', how='left')
    
    with st.spinner("🧮 Расчёт unit-экономики..."):
        # ИСПРАВЛЕНИЕ: Передаем DataFrame и dict напрямую, без сериализации в JSON
        tariffs_map = {row['category']: {k: v for k, v in row.items() if k != 'category'}
                       for _, row in tariff_df.iterrows()}
        
        current_hash = make_hash(df_merged) + tax_label + scheme_label + payment_freq + make_hash(tariffs_map)
        if st.session_state.last_hash == current_hash and not st.session_state.calc_df.empty:
            return st.session_state.calc_df
        
        calc_df = run_calculations_cached(df_merged, tax_label, scheme_label, payment_freq, tariffs_map)
        st.session_state.last_hash = current_hash
        st.session_state.calc_df = calc_df
        return calc_df

def page_dashboard(calc_df: pd.DataFrame):
    st.title("📊 Панель комплексной аналитики")
    if calc_df.empty:
        st.warning("Загрузите данные на вкладке 💾 Импорт / Экспорт.")
        return
    
    c1, c2, c3, c4, c5 = st.columns(5)
    total_revenue = calc_df['selling_price'].sum()
    total_profit = calc_df['gross_profit'].sum()
    avg_margin = calc_df['margin_percent'].mean()
    profitable = (calc_df['gross_profit'] > 0).sum()
    unprofitable = (calc_df['gross_profit'] <= 0).sum()
    
    c1.metric("SKU", len(calc_df))
    c2.metric("Ср. маржа", f"{avg_margin:.2f}%")
    c3.metric("Выручка", format_number(total_revenue, " ₽"))
    c4.metric("Прибыль", format_number(total_profit, " ₽"),
              delta=f"{avg_margin:.1f}% маржа")
    c5.metric("Прибыльных / Убыточных", f"{profitable} / {unprofitable}")
    
    st.markdown("---")
    
    col_left, col_right = st.columns(2)
    with col_left:
        fig = px.treemap(
            calc_df, path=['abc_xyz'], values='gross_profit',
            title="Прибыль по ABC-XYZ сегментам",
            color='margin_percent', color_continuous_scale='RdYlGn',
            color_continuous_midpoint=0
        )
        st.plotly_chart(fig, use_container_width=True)
    
    with col_right:
        fig2 = px.scatter(
            calc_df, x='selling_price', y='gross_profit',
            color='profitability_status', size='daily_sales',
            hover_data=['artikul', 'category', 'margin_percent'],
            title="Карта SKU: Цена vs Прибыль",
            color_discrete_map={
                'Высокомаржинальный': '#2E7D32',
                'Низкомаржинальный': '#F9A825',
                'Убыточный': '#C62828'
            }
        )
        fig2.add_hline(y=0, line_dash="dash", line_color="red")
        st.plotly_chart(fig2, use_container_width=True)
    
    st.subheader("🏆 Топ-10 по прибыли")
    st.dataframe(
        calc_df.nlargest(10, 'gross_profit')[['artikul', 'category', 'selling_price', 'gross_profit', 'margin_percent', 'abc_xyz']],
        use_container_width=True, hide_index=True
    )
    st.subheader("⚠️ Топ-10 убыточных")
    st.dataframe(
        calc_df.nsmallest(10, 'gross_profit')[['artikul', 'category', 'selling_price', 'gross_profit', 'margin_percent']],
        use_container_width=True, hide_index=True
    )

def page_metrics(calc_df: pd.DataFrame, tax_label: str, scheme_label: str,
                 payment_freq: str, tm: HybridTariffManager):
    st.title("🔥 Детальные метрики")
    if calc_df.empty:
        st.warning("Нет данных для отображения.")
        return
    
    with st.expander("🔍 Фильтры", expanded=False):
        col1, col2, col3 = st.columns(3)
        with col1:
            status_filter = st.multiselect(
                "Статус прибыльности",
                options=calc_df['profitability_status'].unique(),
                default=list(calc_df['profitability_status'].unique())
            )
        with col2:
            abc_filter = st.multiselect(
                "ABC-XYZ",
                options=sorted(calc_df['abc_xyz'].unique()),
                default=sorted(calc_df['abc_xyz'].unique())
            )
        with col3:
            cat_filter = st.multiselect(
                "Категория",
                options=sorted(calc_df['category'].unique()),
                default=list(calc_df['category'].unique())
            )
        
        filtered = calc_df[
            calc_df['profitability_status'].isin(status_filter) &
            calc_df['abc_xyz'].isin(abc_filter) &
            calc_df['category'].isin(cat_filter)
        ]
    
    st.dataframe(filtered, use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("💾 Экспорт")
    if OPENPYXL_AVAILABLE:
        excel_data = UltimateExcelExporter.export_max_info(
            filtered, tax_label, scheme_label, tm, payment_freq
        )
        st.download_button(
            label="⬇️ СКАЧАТЬ ПОЛНЫЙ ОТЧЁТ С ЖИВЫМИ ФОРМУЛАМИ (.XLSX)",
            data=excel_data,
            file_name=f"YM_UnitEconomics_Live_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
            use_container_width=True,
            type="primary"
        )
        st.caption("Excel-файл содержит живые формулы VLOOKUP, MIN/MAX, IF. Совместим с Excel 2010+ и Google Sheets. Для датасетов > 1500 строк формулы заменяются значениями во избежание сбоев.")
    else:
        st.error("Установите openpyxl: `pip install openpyxl`")

def page_prices(calc_df: pd.DataFrame):
    st.title("💰 Рекомендованные цены")
    if calc_df.empty:
        st.warning("Нет данных.")
        return
    
    cols = ['artikul', 'category', 'selling_price', 'cogs', 'gross_profit',
            'margin_percent', 'rec_price_min', 'rec_price_15', 'rec_price_25',
            'profitability_status', 'is_price_calc_possible']
    display_cols = [c for c in cols if c in calc_df.columns]
    
    def highlight_prices(row):
        styles = [''] * len(row)
        if row.get('is_price_calc_possible') == False:
            styles[row.index.get_loc('rec_price_min')] = 'background-color: #FFCDD2'
            styles[row.index.get_loc('rec_price_15')] = 'background-color: #FFCDD2'
            styles[row.index.get_loc('rec_price_25')] = 'background-color: #FFCDD2'
        elif row['selling_price'] < row.get('rec_price_min', float('inf')):
            if 'selling_price' in row.index:
                styles[row.index.get_loc('selling_price')] = 'background-color: #FFCDD2'
        return styles
    
    styled = calc_df[display_cols].style.apply(highlight_prices, axis=1)
    st.dataframe(styled, use_container_width=True, hide_index=True)
    
    st.info("🔴 Красным подсвечены SKU, где текущая цена ниже минимальной рентабельной, либо расчет невозможен из-за фатального превышения переменных издержек над ценой.")

def page_tariffs(tm: HybridTariffManager):
    st.title("🗂️ Управление тарифами")
    st.info("Приоритет: API Яндекс Маркета → Загруженный файл → Базовый фоллбэк (15% с предупреждением).")
    
    st.subheader("📋 Текущие тарифы в памяти")
    st.dataframe(tm.to_dataframe(), use_container_width=True, hide_index=True)
    
    st.markdown("---")
    st.subheader("📤 Загрузить справочник тарифов (CSV/XLSX)")
    st.caption("Обязательные столбцы: `category`, `commission_rate`. Опциональные: `sorting_cost`, `delivery_rate`, `delivery_min`, `delivery_max`, `acquiring_transfer_rate`, `acquiring_sku_cost`, `return_rate`, `return_processing`, `storage_fee_per_day`, `special_tariff_rate`, `scheme`")
    
    tariff_file = st.file_uploader("Файл тарифов", type=['csv', 'xlsx'], key="tariff_uploader")
    if tariff_file is not None:
        try:
            t_df_raw = UniversalDataNormalizer.load_file(io.BytesIO(tariff_file.getvalue()), tariff_file.name)
            tm.load_tariffs_from_file(t_df_raw)
            st.success(f"✅ Загружено {len(tm.tariffs)} тарифов. Пересчёт будет выполнен автоматически.")
            st.rerun()
        except Exception as e:
            st.error(f"Ошибка загрузки: {e}")
    
    st.markdown("---")
    if st.button("🗑️ Очистить кэш тарифов"):
        st.session_state.ym_api_cache = {}
        st.success("Кэш очищен")
        st.rerun()

def page_import_export():
    st.title("💾 Центр импорта и редактирования")
    
    uploaded_file = st.file_uploader("Перетащите файл с товарами (CSV/XLSX)", type=['csv', 'xlsx'])
    if uploaded_file is not None:
        try:
            raw_data = UniversalDataNormalizer.load_file(
                io.BytesIO(uploaded_file.getvalue()), uploaded_file.name
            )
            processed_df = UniversalDataNormalizer.normalize_dataframe(raw_data)
            st.session_state.main_df = processed_df
            st.success(f"✅ Импортировано {len(processed_df)} SKU (дубликаты удалены).")
        except Exception as e:
            st.error(f"Ошибка импорта: {str(e)}")
    
    if not st.session_state.main_df.empty:
        st.markdown("---")
        st.subheader("✏️ Редактор данных")
        st.caption("Двойной клик по ячейке для редактирования. Изменения сохраняются автоматически.")
        edited_df = st.data_editor(
            st.session_state.main_df,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="data_editor"
        )
        if not edited_df.equals(st.session_state.main_df):
            st.session_state.main_df = edited_df
            st.session_state.last_hash = ""  # сброс кэша
            st.toast("Данные обновлены! Перейдите на Дашборд для пересчёта.")
        
        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                "⬇️ Скачать нормализованные данные (CSV)",
                st.session_state.main_df.to_csv(index=False).encode('utf-8'),
                f"ym_data_{datetime.now().strftime('%d_%m_%Y')}.csv",
                mime="text/csv"
            )
        with col2:
            if st.button("🗑️ Очистить все данные"):
                st.session_state.main_df = pd.DataFrame()
                st.session_state.calc_df = pd.DataFrame()
                st.session_state.last_hash = ""
                st.rerun()
    else:
        st.info("Загрузите файл или используйте редактор для ручного ввода.")
        
        template = pd.DataFrame({
            'artikul': ['SKU-001', 'SKU-002'],
            'category': ['электроника', 'одежда'],
            'selling_price': [1500, 2500],
            'cogs': [800, 1200],
            'weight_kg': [0.5, 0.3],
            'length_cm': [20, 15],
            'width_cm': [15, 10],
            'height_cm': [5, 3],
            'daily_sales': [5, 12],
            'packaging_cost': [15, 10],
            'first_mile_cost': [50, 40],
            'marketing_budget_per_unit': [30, 50],
            'stock_depth_days': [30, 45],
        })
        st.download_button(
            "📥 Скачать шаблон (CSV)",
            template.to_csv(index=False).encode('utf-8'),
            "ym_template.csv",
            mime="text/csv"
        )

def page_scenario(calc_df: pd.DataFrame):
    st.title("🧪 Сценарный анализ")
    if calc_df.empty:
        st.warning("Сначала загрузите данные и выполните расчёт.")
        return
    
    st.subheader("Что будет, если изменить цену / себестоимость / рекламу?")
    
    col1, col2, col3 = st.columns(3)
    with col1:
        price_change = st.slider("Изменение цены, %", -30, 30, 0, 1)
    with col2:
        cogs_change = st.slider("Изменение себестоимости, %", -30, 30, 0, 1)
    with col3:
        marketing_change = st.slider("Изменение рекламы, %", -50, 100, 0, 5)
    
    scenario = calc_df.copy()
    scenario['selling_price'] = scenario['selling_price'] * (1 + price_change / 100)
    scenario['cogs'] = scenario['cogs'] * (1 + cogs_change / 100)
    scenario['marketing_budget_per_unit'] = scenario['marketing_budget_per_unit'] * (1 + marketing_change / 100)
    
    scenario['gross_profit_scenario'] = (
        scenario['gross_profit'] +
        (scenario['selling_price'] - calc_df['selling_price']) -
        (scenario['cogs'] - calc_df['cogs']) -
        (scenario['marketing_budget_per_unit'] - calc_df['marketing_budget_per_unit'])
    )
    scenario['margin_scenario'] = np.where(
        scenario['selling_price'] > 0,
        scenario['gross_profit_scenario'] / scenario['selling_price'] * 100,
        0
    )
    
    total_before = calc_df['gross_profit'].sum()
    total_after = scenario['gross_profit_scenario'].sum()
    
    st.metric("Общая прибыль (было → стало)",
              f"{format_number(total_after, ' ₽')}",
              delta=f"{total_after - total_before:,.0f} ₽")
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=['Было', 'Стало'],
        y=[total_before, total_after],
        marker_color=['#90A4AE', '#43A047']
    ))
    fig.update_layout(title="Сравнение прибыли", yaxis_title="₽")
    st.plotly_chart(fig, use_container_width=True)
    
    st.dataframe(
        scenario[['artikul', 'category', 'selling_price', 'gross_profit_scenario', 'margin_scenario']].rename(
            columns={'gross_profit_scenario': 'Прибыль (сценарий)', 'margin_scenario': 'Маржа % (сценарий)'}
        ),
        use_container_width=True, hide_index=True
    )

def main():
    st.set_page_config(
        page_title=APP_NAME, page_icon="📈", layout="wide",
        initial_sidebar_state="expanded"
    )
    init_session_state()
    
    scheme_label, tax_label, payment_freq, use_api, ym_api = render_sidebar()
    
    page = st.sidebar.radio("Навигация:", [
        "📊 Дашборд", "🔥 Метрики и ABC-XYZ", "💰 Рекомендованные цены",
        "🗂️ Тарифы и Справочник", "💾 Импорт / Экспорт", "🧪 Сценарный анализ"
    ])
    
    tm = HybridTariffManager()
    
    if not st.session_state.main_df.empty:
        calc_df = run_analysis(
            st.session_state.main_df, tax_label, scheme_label,
            payment_freq, tm, ym_api, use_api
        )
    else:
        calc_df = pd.DataFrame()
    
    if page == "📊 Дашборд":
        page_dashboard(calc_df)
    elif page == "🔥 Метрики и ABC-XYZ":
        page_metrics(calc_df, tax_label, scheme_label, payment_freq, tm)
    elif page == "💰 Рекомендованные цены":
        page_prices(calc_df)
    elif page == "🗂️ Тарифы и Справочник":
        page_tariffs(tm)
    elif page == "💾 Импорт / Экспорт":
        page_import_export()
    elif page == "🧪 Сценарный анализ":
        page_scenario(calc_df)

if __name__ == "__main__":
    main()
