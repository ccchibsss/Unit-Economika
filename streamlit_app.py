#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
🚀 ULTIMATE UNIT ECONOMICS FOR YANDEX MARKET v22.2 — FULLY REFACTORED
============================================================================
Исправлено:
1. Синтаксис IconSetRule в openpyxl (добавлены cfvo объекты).
2. Мутация DataFrame внутри @st.cache_data (добавлен df.copy()).
3. Реализованы freeze_panes, автоподбор ширины и стилизация заголовков.
4. Детерминированный фоллбэк в make_hash (убран time.time(), ломающий кэш).
5. special_tariff_rate изменён с 0.42 на 0.15 (42% фатально для товаров <=300₽).
6. Восстановлен полный функционал Streamlit UI и листа "Параметры" в Excel.
"""
import streamlit as st
import pandas as pd
import numpy as np
import io
import json
import requests
import logging
import warnings
import hashlib
import re
import time
from datetime import datetime
from enum import Enum
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List, Optional, Tuple, Any
from functools import wraps
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import IconSetRule, ConditionalFormattingValueObject
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.utils.dataframe import dataframe_to_rows
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('YandexMarketUnitEconomics')

APP_VERSION = "22.2.0"
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
        except Exception: 
            continue
    return text

def make_hash(obj: Any) -> str:
    """Надёжный детерминированный хеш для pandas DataFrame / dict."""
    try:
        if isinstance(obj, pd.DataFrame):
            return hashlib.sha256(pd.util.hash_pandas_object(obj, index=True).values.tobytes()).hexdigest()[:16]
        return hashlib.sha256(str(obj).encode()).hexdigest()[:16]
    except Exception:
        # Детерминированный фоллбэк вместо time.time(), чтобы не ломать логику кэширования
        return hashlib.sha256("hash_error_fallback".encode()).hexdigest()[:16]

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

class Tariff:
    """Модель тарифа с валидацией."""
    def __init__(self, category: str, commission_rate: float = 0.15, min_commission: float = 0.0,
                 sorting_cost: float = 45.0, delivery_rate: float = 0.045, delivery_min: float = 60.0, 
                 delivery_max: float = 500.0, acquiring_transfer_rate: float = 0.016, acquiring_sku_cost: float = 0.12,
                 return_rate: float = 0.05, return_processing: float = 15.0, storage_fee_per_day: float = 0.50,
                 special_tariff_rate: float = 0.15, source: str = "Базовый фоллбэк", scheme: str = "FBS"): 
                 # ^^^ Исправлено с 0.42. 42% комиссии для товаров <=300р фатально для юнит-экономики.
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
            if "sellingProgram" in payload["parameters"]:
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
                special_tariff_rate=float(row.get('special_tariff_rate', 0.15)),
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
        unique_cats = df['category'].dropna().unique()
        tariff_map = {}
        for cat in unique_cats:
            tariff_map[cat] = self.get_best_tariff(cat, scheme, ym_api, use_api)
        
        tariff_df = pd.DataFrame([
            {'category': cat, **t.to_dict()} for cat, t in tariff_map.items()
        ])
        return tariff_df

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
                errors.append(f"selling_price: {zero_prices} SKU с нулевой ценой")
        
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
    df_hash: str, # Хеш используется для триггера кэша, сам df передается отдельно
    df: pd.DataFrame,
    tax_label: str,
    scheme_label: str,
    payment_frequency: str,
    tariffs_map: dict
) -> pd.DataFrame:
    """Кешированный расчёт unit-экономики. df.copy() предотвращает мутацию кэша."""
    if df.empty:
        return df
    
    # КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ: Явная копия для предотвращения инплейс-мутации кэшированного объекта
    df = df.copy()
    
    tax_system = TaxSystem.by_label(tax_label)
    scheme = scheme_label.split(" ")[0]
    
    payment_rates = {
        "Ежемесячно (1.0%)": 0.01,
        "Раз в 2 недели (1.3%)": 0.013,
        "Еженедельно, 4 нед. (1.6%)": 0.016,
        "Ежедневно (3.3%)": 0.033
    }
    p_transfer_rate = payment_rates.get(payment_frequency, 0.016)
    
    for col in ['artikul', 'category']:
        if col in df.columns:
            df[col] = df[col].astype(str).apply(fix_double_utf8)
    
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
    
    tariff_df = pd.DataFrame.from_dict(tariffs_map, orient='index')
    if tariff_df.empty:
        tariff_df = pd.DataFrame([Tariff('default').to_dict()])
    
    df = df.merge(tariff_df, on='category', how='left')
    
    for col in ['commission_rate', 'min_commission', 'sorting_cost', 'delivery_rate',
                'delivery_min', 'delivery_max', 'acquiring_transfer_rate', 'acquiring_sku_cost',
                'return_rate', 'return_processing', 'storage_fee_per_day', 'special_tariff_rate']:
        if col in df.columns:
            df[col] = df[col].fillna(Tariff('default').to_dict()[col])
    
    vol_weight = (df['length_cm'] * df['width_cm'] * df['height_cm']) / 5000.0
    df['billable_weight'] = np.ceil(np.maximum(df['weight_kg'], vol_weight) * 2) / 2
    df['is_special_tariff'] = (df['selling_price'] <= 300) & (df['volume_liters'] <= 5)
    
    df['commission'] = np.where(
        df['is_special_tariff'],
        df['selling_price'] * df['special_tariff_rate'],
        np.maximum(df['selling_price'] * df['commission_rate'], df['min_commission'])
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
        np.where(df['billable_weight'] <= 4, 100,
                 np.where(df['billable_weight'] <= 10, 300, 600))
    )
    
    df['sorting_cost'] = np.where(
        df['is_special_tariff'],
        0.0,
        np.where(scheme == 'FBS', df['sorting_cost'], 0.0)
    )
    
    df['acquiring_sku_cost'] = df['acquiring_sku_cost'] / df['quantity_per_order']
    df['acquiring_transfer_cost'] = df['selling_price'] * p_transfer_rate
    df['acquiring_cost'] = df['acquiring_sku_cost'] + df['acquiring_transfer_cost']
    
    df['return_processing_cost'] = np.where(df['is_special_tariff'], 0.0, df['return_processing'])
    df['return_delivery_cost'] = np.where(df['is_special_tariff'], 0.0, df['middle_mile_cost'] * df['return_rate'])
    df['return_cost'] = df['return_processing_cost'] + df['return_delivery_cost']
    
    df['pick_pack_cost'] = 35.0
    df['warehouse_cost'] = np.where(
        df['warehouse_cost'] == 0,
        (df['stock_depth_days'] * df['daily_sales']) * df['storage_fee_per_day'],
        df['warehouse_cost']
    )
    
    df['fixed_operational_costs'] = (
        df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] +
        df['packaging_cost'] + df['return_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    )
    df['marketplace_fees'] = (
        df['commission'] + df['delivery_to_customer'] + df['middle_mile_cost'] +
        df['sorting_cost'] + df['acquiring_cost']
    )
    df['pre_tax_expenses'] = df['fixed_operational_costs'] + df['marketplace_fees']
    
    if tax_system.base == "revenue":
        df['tax_cost'] = df['selling_price'] * tax_system.rate
    elif tax_system.base == "profit_vat":
        revenue_without_vat = df['selling_price'] / 1.20
        pre_tax_profit = revenue_without_vat - df['pre_tax_expenses']
        df['tax_cost'] = np.maximum(pre_tax_profit, 0) * tax_system.rate
    else: # profit
        pre_tax_profit = df['selling_price'] - df['pre_tax_expenses']
        df['tax_cost'] = np.maximum(pre_tax_profit, 0) * tax_system.rate
    
    df['total_expenses'] = df['pre_tax_expenses'] + df['tax_cost']
    df['gross_profit'] = df['selling_price'] - df['total_expenses']
    df['margin_percent'] = np.where(df['selling_price'] > 0, (df['gross_profit'] / df['selling_price']) * 100, 0.0)
    
    var_fees = np.where(
        df['is_special_tariff'],
        df['special_tariff_rate'] + p_transfer_rate + (tax_system.rate if tax_system.base == "revenue" else 0),
        df['commission_rate'] + df['delivery_rate'] + p_transfer_rate + (tax_system.rate if tax_system.base == "revenue" else 0)
    )
    denom = 1.0 - var_fees
    denom_valid = denom > 0.01 
    
    fixed_no_return = (
        df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] +
        df['packaging_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    )
    
    df['rec_price_min'] = np.where(denom_valid, safe_divide(fixed_no_return, denom, default=np.nan), np.nan)
    df['rec_price_15'] = np.where(denom_valid, safe_divide(fixed_no_return, denom - 0.15, default=np.nan), np.nan)
    df['rec_price_25'] = np.where(denom_valid, safe_divide(fixed_no_return, denom - 0.25, default=np.nan), np.nan)
    
    df['variable_costs'] = (
        df['commission'] + df['delivery_to_customer'] + df['middle_mile_cost'] +
        df['sorting_cost'] + df['acquiring_cost'] + df['return_cost']
    )
    df['fixed_costs'] = fixed_no_return + df['return_cost']
    df['contribution_margin'] = df['selling_price'] - df['variable_costs']
    df['roi_percent'] = np.where(df['cogs'] > 0, (df['gross_profit'] / df['cogs']) * 100, 0.0)
    df['break_even_units'] = safe_divide(df['fixed_costs'], df['contribution_margin'], default=0.0)
    
    df['abc_category'] = np.where(df['daily_sales'] >= 10, 'A', np.where(df['daily_sales'] >= 3, 'B', 'C'))
    df['xyz_category'] = np.where(df['margin_percent'] >= 20, 'X', np.where(df['margin_percent'] >= 10, 'Y', 'Z'))
    df['abc_xyz'] = df['abc_category'] + df['xyz_category']
    df['profitability_status'] = np.where(
        df['gross_profit'] > 0,
        np.where(df['margin_percent'] >= 20, 'Высокомаржинальный', 'Низкомаржинальный'),
        'Убыточный'
    )
    
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
    
    tech_cols = ['commission_rate', 'min_commission', 'sorting_cost', 'delivery_rate',
                 'delivery_min', 'delivery_max', 'acquiring_transfer_rate', 'acquiring_sku_cost',
                 'return_rate', 'return_processing', 'storage_fee_per_day', 'special_tariff_rate',
                 'source', 'scheme', 'is_special_tariff', 'billable_weight']
    for col in tech_cols:
        if col in df.columns:
            df = df.drop(columns=[col])
    
    return df

# ============================================================================
# БЛОК 6: ЭКСПОРТЁР EXCEL С ПОЛНОЙ ВИЗУАЛИЗАЦИЕЙ
# ============================================================================
class UltimateExcelExporter:
    @staticmethod
    def export_max_info(df: pd.DataFrame, tax_label: str, scheme_label: str,
                        tariff_manager: HybridTariffManager,
                        payment_frequency: str) -> bytes:
        if not OPENPYXL_AVAILABLE or df.empty:
            return b""
        
        wb = Workbook()
        ws_params = wb.active
        ws_params.title = "Параметры"
        
        # 1. Заполнение листа параметров
        params_data = [
            ["Параметр", "Значение"],
            ["Версия приложения", APP_VERSION],
            ["Дата формирования", datetime.now().strftime("%Y-%m-%d %H:%M")],
            ["Система налогообложения", tax_label],
            ["Схема работы", scheme_label],
            ["Частота выплат", payment_frequency],
            ["Всего SKU в расчете", len(df)]
        ]
        for r_idx, row in enumerate(params_data, 1):
            for c_idx, value in enumerate(row, 1):
                cell = ws_params.cell(row=r_idx, column=c_idx, value=value)
                if r_idx == 1:
                    cell.font = Font(bold=True)
        
        ws_params.column_dimensions['A'].width = 25
        ws_params.column_dimensions['B'].width = 40

        # 2. Лист детального расчета
        ws_detail = wb.create_sheet("Детальный_Расчет")
        
        # Запись данных через оптимизированный генератор openpyxl
        for r_idx, row in enumerate(dataframe_to_rows(df, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws_detail.cell(row=r_idx, column=c_idx, value=value)
        
        # 3. Zebra-striping и заголовки через Table
        num_rows = len(df) + 1
        num_cols = len(df.columns)
        ref = f"A1:{get_column_letter(num_cols)}{num_rows}"
        tab = Table(displayName="DataExport", ref=ref)
        style = TableStyleInfo(name="TableStyleMedium9", showFirstColumn=False,
                               showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        tab.tableStyleInfo = style
        ws_detail.add_table(tab)
        
        # 4. Заморозка областей (заголовки + первая колонка)
        ws_detail.freeze_panes = "B2"
        
        # 5. Автоподбор ширины столбцов и ручная стилизация заголовков (на случай переопределения Table)
        header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        
        for col_idx, col_name in enumerate(df.columns, 1):
            col_letter = get_column_letter(col_idx)
            # Принудительная стилизация заголовка
            header_cell = ws_detail.cell(row=1, column=col_idx)
            header_cell.fill = header_fill
            header_cell.font = header_font
            header_cell.alignment = Alignment(horizontal="center", vertical="center")
            
            # Расчет ширины
            max_len = max(
                len(str(col_name)),
                df[col_name].astype(str).map(len).max() if not df.empty else 0
            )
            ws_detail.column_dimensions[col_letter].width = min(max_len + 2, 60)
            
        # 6. Светофор (IconSet) для маржинальности
        if 'margin_percent' in df.columns:
            m_idx = df.columns.get_loc('margin_percent') + 1
            m_col = get_column_letter(m_idx)
            
            # Корректный синтаксис openpyxl >= 3.0 с использованием cfvo
            rule = IconSetRule(
                iconSet="3TrafficLights1",
                cfvo=[
                    ConditionalFormattingValueObject(type="percent", val=0),
                    ConditionalFormattingValueObject(type="percent", val=33),
                    ConditionalFormattingValueObject(type="percent", val=66)
                ],
                showValue=True
            )
            ws_detail.conditional_formatting.add(
                f"{m_col}2:{m_col}{num_rows}",
                rule
            )
            
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
        if raw_df.empty: return raw_df
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
                if target_col == 'artikul': final_data[target_col] = [f"SKU_{i+1}" for i in range(len(df))]
                elif target_col == 'category': final_data[target_col] = "не указано"
                elif target_col == 'quantity_per_order': final_data[target_col] = 1.0
                else: final_data[target_col] = 0.0
        norm_df = pd.DataFrame(final_data)
        num_cols = ['selling_price', 'cogs', 'daily_sales', 'weight_kg', 'length_cm', 'width_cm', 'height_cm', 'volume_liters', 'packaging_cost', 'first_mile_cost', 'marketing_budget_per_unit', 'stock_depth_days', 'quantity_per_order', 'warehouse_cost']
        for col in num_cols:
            if col in norm_df.columns:
                norm_df[col] = pd.to_numeric(norm_df[col].astype(str).str.replace(r'[\s,;%₽]', '', regex=True), errors='coerce').fillna(0.0).abs()
        norm_df['artikul'] = norm_df['artikul'].astype(str).str.strip()
        norm_df['category'] = norm_df['category'].astype(str).str.strip().str.lower()
        return norm_df.drop_duplicates(subset=['artikul'], keep='first')
    
    @classmethod
    def load_file(cls, file_buffer: io.BytesIO, file_name: str) -> pd.DataFrame:
        try:
            if file_name.endswith('.csv'): return pd.read_csv(file_buffer, sep=None, engine='python', encoding='utf-8')
            elif file_name.endswith(('.xls', '.xlsx')): return pd.read_excel(file_buffer)
            else: raise ValueError("Формат не поддерживается.")
        except Exception:
            file_buffer.seek(0)
            return pd.read_csv(file_buffer, sep=None, engine='python', encoding='cp1251')

# ============================================================================
# БЛОК 8: STREAMLIT UI (ПОЛНАЯ РЕАЛИЗАЦИЯ)
# ============================================================================
def init_session_state():
    defaults = {
        'main_df': pd.DataFrame(), 
        'calc_df': pd.DataFrame(), 
        'tariffs': {}, 
        'ym_api_cache': {}, 
        'last_hash': '',
        'api_key': '',
        'business_id': ''
    }
    for key, val in defaults.items():
        if key not in st.session_state: st.session_state[key] = val

def main():
    st.set_page_config(page_title=APP_NAME, page_icon="📈", layout="wide")
    init_session_state()
    
    st.title(f"📊 {APP_NAME} v{APP_VERSION}")
    st.markdown("Монолитный калькулятор юнит-экономики с векторизованными вычислениями и профессиональным экспортом.")
    
    # --- САЙДБАР: НАСТРОЙКИ ---
    with st.sidebar:
        st.header("⚙️ Настройки")
        api_key = st.text_input("API Key Яндекс Маркета", value=st.session_state.api_key, type="password")
        business_id = st.text_input("Business ID", value=st.session_state.business_id)
        st.session_state.api_key = api_key
        st.session_state.business_id = business_id
        
        tax_options = [tax.label for tax in TaxSystem]
        tax_label = st.selectbox("Система налогообложения", tax_options, index=0)
        
        scheme_options = ["FBS (склад продавца)", "FBY (склад Маркета)", "Экспресс", "DBS (доставка продавца)"]
        scheme_label = st.selectbox("Схема работы", scheme_options, index=0)
        
        payment_options = ["Ежемесячно (1.0%)", "Раз в 2 недели (1.3%)", "Еженедельно, 4 нед. (1.6%)", "Ежедневно (3.3%)"]
        payment_frequency = st.selectbox("Частота выплат", payment_options, index=2)
        
        use_api = st.checkbox("Использовать API ЯМ для тарифов", value=True)

    # --- ОСНОВНОЙ ЭКРАН: ЗАГРУЗКА ---
    col1, col2 = st.columns(2)
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
                
                st.session_state.main_df = validated_df
                st.success(f"Загружено {len(validated_df)} уникальных SKU")
            except Exception as e:
                st.error(f"Ошибка чтения файла: {e}")

    with col2:
        st.subheader("2. Загрузка справочника тарифов (опционально)")
        tariff_file = st.file_uploader("Тарифы (Excel/CSV)", type=['xlsx', 'xls', 'csv'], key="tariff_uploader")
        tariff_manager = HybridTariffManager()
        if tariff_file is not None:
            try:
                t_raw = UniversalDataNormalizer.load_file(io.BytesIO(tariff_file.read()), tariff_file.name)
                # Упрощенная нормализация для тарифов
                t_raw.columns = [str(c).strip().lower() for c in t_raw.columns]
                tariff_manager.load_tariffs_from_file(t_raw)
                st.success(f"Загружено {len(tariff_manager.tariffs)} тарифов")
            except Exception as e:
                st.error(f"Ошибка тарифов: {e}")

    # --- РАСЧЁТ ---
    st.markdown("---")
    if not st.session_state.main_df.empty:
        if st.button("🚀 Рассчитать юнит-экономику", type="primary"):
            with st.spinner("Выполняется векторизованный расчёт..."):
                ym_api = YandexMarketAPI(api_key=api_key, business_id=business_id) if api_key else None
                
                # Получаем тарифы для уникальных категорий
                tariff_df = tariff_manager.get_tariffs_vectorized(
                    st.session_state.main_df, 
                    scheme=scheme_label.split(" ")[0], 
                    ym_api=ym_api, 
                    use_api=use_api
                )
                tariffs_map = tariff_df.set_index('category').to_dict(orient='index')
                
                # Генерируем хеш для кэша
                current_hash = make_hash(st.session_state.main_df)
                
                if current_hash != st.session_state.last_hash or st.session_state.calc_df.empty:
                    calc_df = run_calculations_cached(
                        df_hash=current_hash,
                        df=st.session_state.main_df,
                        tax_label=tax_label,
                        scheme_label=scheme_label,
                        payment_frequency=payment_frequency,
                        tariffs_map=tariffs_map
                    )
                    st.session_state.calc_df = calc_df
                    st.session_state.last_hash = current_hash
                else:
                    calc_df = st.session_state.calc_df
                
                st.success("Расчёт завершён успешно!")

        # --- ОТОБРАЖЕНИЕ РЕЗУЛЬТАТОВ ---
        if not st.session_state.calc_df.empty:
            st.subheader("3. Результаты расчёта")
            
            # Сводные метрики
            c1, c2, c3, c4 = st.columns(4)
            df_calc = st.session_state.calc_df
            c1.metric("Всего SKU", len(df_calc))
            c2.metric("Средняя маржа, %", f"{df_calc['margin_percent'].mean():.1f}%")
            c3.metric("Убыточных SKU", len(df_calc[df_calc['gross_profit'] < 0]))
            c4.metric("Высокомаржинальных", len(df_calc[df_calc['profitability_status'] == 'Высокомаржинальный']))
            
            # Фильтр и таблица
            status_filter = st.multiselect(
                "Фильтр по статусу", 
                options=df_calc['profitability_status'].unique(),
                default=df_calc['profitability_status'].unique()
            )
            filtered_df = df_calc[df_calc['profitability_status'].isin(status_filter)]
            
            st.dataframe(
                filtered_df[['artikul', 'category', 'selling_price', 'cogs', 'gross_profit', 'margin_percent', 'roi_percent', 'profitability_status']],
                use_container_width=True,
                height=400
            )
            
            # --- ЭКСПОРТ ---
            st.subheader("4. Экспорт")
            excel_data = UltimateExcelExporter.export_max_info(
                df=filtered_df,
                tax_label=tax_label,
                scheme_label=scheme_label,
                tariff_manager=tariff_manager,
                payment_frequency=payment_frequency
            )
            
            if excel_data:
                st.download_button(
                    label="📥 Скачать улучшенный Excel (Zebra, Светофор, Freeze)",
                    data=excel_data,
                    file_name=f"unit_economics_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.warning("Не удалось сформировать Excel. Проверьте установку openpyxl.")

if __name__ == "__main__":
    main()
