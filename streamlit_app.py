#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
🚀 ULTIMATE UNIT ECONOMICS FOR YANDEX MARKET v20.0 - MAX INFO EXPORT
============================================================================
Максимально точные подсчёты для Яндекс Маркета с живыми формулами в Excel.
Тарифы: API Яндекс Маркета + DeepSeek fallback + гибридный метод.
Актуальные тарифы на август 2026.
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
from enum import Enum
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List, Tuple, Optional

warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('YandexMarketUnitEconomics')

OPENPYXL_AVAILABLE = False
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import DataBarRule, ColorScaleRule, FormulaRule
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    pass

APP_VERSION = "20.0.0"
APP_NAME = "Yandex Market Unit Economics PRO"

# ============================================================================
# БЛОК 0: СЛУЖЕБНЫЕ УТИЛИТЫ
# ============================================================================
def money_round(value: float) -> float:
    if pd.isna(value) or np.isinf(value):
        return 0.0
    return float(Decimal(str(value)).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP))

def percent_round(value: float) -> float:
    if pd.isna(value) or np.isinf(value):
        return 0.0
    return float(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def fix_double_utf8(text: str) -> str:
    if not isinstance(text, str) or not text:
        return text
    for source_enc, target_enc in [('cp1251', 'utf-8'), ('latin1', 'utf-8')]:
        try:
            fixed = text.encode(source_enc).decode(target_enc)
            if fixed and 'Р' not in fixed[:2]:
                return fixed
        except Exception:
            continue
    return text

def format_number(num: float, suffix='') -> str:
    if pd.isna(num):
        return "0"
    abs_num = abs(num)
    sign = "-" if num < 0 else ""
    for unit in ['', 'K', 'M', 'B']:
        if abs_num < 1000.0:
            return f"{sign}{abs_num:3.1f}{unit}{suffix}".strip()
        abs_num /= 1000.0
    return f"{sign}{abs_num:.1f}T{suffix}"

# ============================================================================
# БЛОК 1: КОНФИГУРАЦИИ
# ============================================================================
class TaxSystem(Enum):
    USN_6 = ("УСН 6% (доходы)", 0.06, "revenue", 0.0)
    USN_15 = ("УСН 15% (доходы-расходы)", 0.15, "profit", 0.01)
    OSN = ("ОСН (общая с НДС 20%)", 0.20, "profit_vat", 0.0)
    AUSN_8 = ("АУСН 8% (доходы)", 0.08, "revenue", 0.0)
    AUSN_20 = ("АУСН 20% (доходы-расходы)", 0.20, "profit", 0.0)
    
    def __init__(self, label, rate, base, min_rate):
        self.label = label
        self.rate = rate
        self.base = base
        self.min_rate = min_rate
    
    @classmethod
    def by_label(cls, label):
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
    def __init__(self, category: str, commission_rate: float = 0.12, min_commission: float = 0.0,
                 sorting_cost: float = 45.0, middle_mile_base: float = 0.0, middle_mile_per_kg: float = 0.0,
                 delivery_rate: float = 0.045, delivery_min: float = 60.0, delivery_max: float = 500.0,
                 acquiring_transfer_rate: float = 0.016, acquiring_sku_cost: float = 0.12,
                 return_rate: float = 0.05, return_processing: float = 15.0,
                 storage_fee_per_day: float = 0.50, special_tariff_threshold: float = 300.0,
                 special_tariff_volume: float = 5.0, special_tariff_rate: float = 0.42,
                 source: str = "Справочник", scheme: str = "FBS"):
        self.category = category
        self.commission_rate = commission_rate
        self.min_commission = min_commission
        self.sorting_cost = sorting_cost
        self.middle_mile_base = middle_mile_base
        self.middle_mile_per_kg = middle_mile_per_kg
        self.delivery_rate = delivery_rate
        self.delivery_min = delivery_min
        self.delivery_max = delivery_max
        self.acquiring_transfer_rate = acquiring_transfer_rate
        self.acquiring_sku_cost = acquiring_sku_cost
        self.return_rate = return_rate
        self.return_processing = return_processing
        self.storage_fee_per_day = storage_fee_per_day
        self.special_tariff_threshold = special_tariff_threshold
        self.special_tariff_volume = special_tariff_volume
        self.special_tariff_rate = special_tariff_rate
        self.source = source
        self.scheme = scheme

# ============================================================================
# БЛОК 2: API + ГИБРИДНЫЙ МЕНЕДЖЕР ТАРИФОВ
# ============================================================================
class YandexMarketAPI:
    BASE_URL = "https://api.partner.market.yandex.ru"
    
    def __init__(self, api_key: str, business_id: str = None):
        self.api_key = api_key
        self.business_id = business_id
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json"
        }
        if business_id:
            self.headers["X-Business-Id"] = business_id
    
    def get_campaigns(self) -> List[Dict]:
        try:
            resp = requests.get(f"{self.BASE_URL}/v2/campaigns", headers=self.headers, timeout=30)
            resp.raise_for_status()
            return resp.json().get("campaigns", [])
        except Exception as e:
            logger.error(f"Ошибка получения кампаний: {e}")
            return []
    
    def calculate_tariffs(self, offers: List[Dict], campaign_id: int = None, 
                          selling_program: str = "FBS", frequency: str = "WEEKLY",
                          payment_delay_weeks: int = 4) -> List[Dict]:
        try:
            payload = {
                "parameters": {
                    "sellingProgram": selling_program,
                    "frequency": frequency,
                    "paymentDelayWeeks": payment_delay_weeks,
                    "currency": "RUR"
                },
                "offers": offers
            }
            if campaign_id:
                payload["parameters"]["campaignId"] = campaign_id
                del payload["parameters"]["sellingProgram"]
            
            resp = requests.post(
                f"{self.BASE_URL}/v2/tariffs/calculate",
                headers=self.headers,
                json=payload,
                timeout=60
            )
            resp.raise_for_status()
            data = resp.json()
            return data.get("result", {}).get("offers", [])
        except Exception as e:
            logger.error(f"Ошибка расчёта тарифов API: {e}")
            return []

class DeepSeekAPI:
    BASE_URL = "https://api.deepseek.com"
    
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.headers = {
            "Authorization": f"Bearer {api_key}",
            "Content-Type": "application/json"
        }
    
    def analyze_tariffs(self, category_name: str, scheme: str = "FBS") -> Dict:
        try:
            prompt = f"""Ты — эксперт по тарифам Яндекс Маркета. 
Для категории товаров "{category_name}" и схемы работы {scheme} 
укажи актуальные тарифы Яндекс Маркета на август 2026 года в формате JSON:
{{
  "commission_rate": 0.10,
  "min_commission": 0,
  "sorting_cost": 45,
  "middle_mile_per_kg": 0,
  "delivery_rate": 0.045,
  "delivery_min": 60,
  "delivery_max": 500,
  "acquiring_transfer_rate": 0.016,
  "acquiring_sku_cost": 0.12,
  "return_rate": 0.05,
  "return_processing": 15,
  "storage_fee_per_day": 0.5,
  "special_tariff_rate": 0.42,
  "source": "DeepSeek AI"
}}
Ответь ТОЛЬКО JSON, без пояснений."""
            
            resp = requests.post(
                f"{self.BASE_URL}/chat/completions",
                headers=self.headers,
                json={
                    "model": "deepseek-v4-flash",
                    "messages": [{"role": "user", "content": prompt}],
                    "temperature": 0.1,
                    "max_tokens": 500
                },
                timeout=30
            )
            resp.raise_for_status()
            content = resp.json()["choices"][0]["message"]["content"]
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            if json_match:
                return json.loads(json_match.group())
            return {}
        except Exception as e:
            logger.error(f"Ошибка DeepSeek API: {e}")
            return {}

class HybridTariffManager:
    DEFAULTS = {
        'default': Tariff('default', 0.12, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Локальная база (август 2026)"),
        'автозапчасти': Tariff('автозапчасти', 0.10, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Оферта ЯМ: Авто 5-19%"),
        'авто': Tariff('авто', 0.10, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Оферта ЯМ: Авто 5-19%"),
        'электроника': Tariff('электроника', 0.10, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Оферта ЯМ: Электроника 10-23%"),
        'одежда': Tariff('одежда', 0.185, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Оферта ЯМ: Одежда 14.5-23%"),
        'обувь': Tariff('обувь', 0.185, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Оферта ЯМ: Обувь 14.5-23%"),
        'аксессуары': Tariff('аксессуары', 0.14, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Оферта ЯМ: Аксессуары 14-23%"),
        'детские товары': Tariff('детские товары', 0.185, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Оферта ЯМ: Детские 5-20%"),
        'товары для дома': Tariff('товары для дома', 0.16, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Оферта ЯМ: Дом 13-19%"),
        'спорт и отдых': Tariff('спорт и отдых', 0.16, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Оферта ЯМ: Спорт 11-21%"),
        'красота': Tariff('красота', 0.125, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Оферта ЯМ: Красота 5-20%"),
        'продукты': Tariff('продукты', 0.115, 0, 45, 0, 0, 0.045, 60, 500, 0.016, 0.12, 0.05, 15, 0.50, 300, 5, 0.42, "Оферта ЯМ: Продукты 8-15%"),
    }
    
    MIDDLE_MILE_TARIFFS = {
        'fbs': [(0, 4, 80, 190), (4, 10, 270, 440), (10, 20, 530, 800), (20, 999, 1000, 3500)],
        'fby': [(0, 4, 60, 160), (4, 10, 230, 370), (10, 20, 440, 670), (20, 999, 840, 3500)],
        'express': [(0, 25, 0, 0)]
    }
    
    def __init__(self):
        if 'tariffs' not in st.session_state:
            st.session_state.tariffs = dict(self.DEFAULTS)
        if 'ym_api_cache' not in st.session_state:
            st.session_state.ym_api_cache = {}
    
    @property
    def tariffs(self):
        return st.session_state.tariffs
    
    def get_middle_mile_cost(self, weight_kg: float, scheme: str = "FBS") -> float:
        scheme_lower = scheme.lower()
        if scheme_lower not in self.MIDDLE_MILE_TARIFFS:
            scheme_lower = 'fbs'
        for min_w, max_w, fbs_rate, fby_rate in self.MIDDLE_MILE_TARIFFS[scheme_lower]:
            if min_w <= weight_kg < max_w:
                return fbs_rate if scheme_lower == 'fbs' else fby_rate
        return 0.0
    
    def get_best_tariff(self, category_name: str, scheme: str = "FBS",
                        ym_api: YandexMarketAPI = None, deepseek_api: DeepSeekAPI = None,
                        use_api: bool = True) -> Tariff:
        cat_clean = str(category_name).lower().strip()
        cache_key = f"{cat_clean}_{scheme}"
        
        if cache_key in st.session_state.ym_api_cache:
            return st.session_state.ym_api_cache[cache_key]
        
        # 1. API Яндекс Маркета
        if use_api and ym_api and ym_api.api_key:
            try:
                test_offers = [{"categoryId": 0, "price": 1000, "length": 10, "width": 10, "height": 10, "weight": 1, "quantity": 1}]
                result = ym_api.calculate_tariffs(test_offers, selling_program=scheme)
                if result and len(result) > 0:
                    tariffs_data = result[0].get("tariffs", [])
                    t = self._parse_ym_tariffs(tariffs_data, cat_clean, scheme)
                    if t:
                        st.session_state.ym_api_cache[cache_key] = t
                        return t
            except Exception as e:
                logger.warning(f"API ЯМ недоступен для {cat_clean}: {e}")
        
        # 2. DeepSeek Fallback
        if use_api and deepseek_api and deepseek_api.api_key:
            try:
                ds_result = deepseek_api.analyze_tariffs(cat_clean, scheme)
                if ds_result:
                    t = Tariff(
                        category=cat_clean,
                        commission_rate=ds_result.get('commission_rate', 0.12),
                        min_commission=ds_result.get('min_commission', 0),
                        sorting_cost=ds_result.get('sorting_cost', 45),
                        delivery_rate=ds_result.get('delivery_rate', 0.045),
                        delivery_min=ds_result.get('delivery_min', 60),
                        delivery_max=ds_result.get('delivery_max', 500),
                        acquiring_transfer_rate=ds_result.get('acquiring_transfer_rate', 0.016),
                        acquiring_sku_cost=ds_result.get('acquiring_sku_cost', 0.12),
                        return_rate=ds_result.get('return_rate', 0.05),
                        return_processing=ds_result.get('return_processing', 15),
                        storage_fee_per_day=ds_result.get('storage_fee_per_day', 0.5),
                        special_tariff_rate=ds_result.get('special_tariff_rate', 0.42),
                        source=f"DeepSeek AI ({scheme})",
                        scheme=scheme
                    )
                    st.session_state.ym_api_cache[cache_key] = t
                    return t
            except Exception as e:
                logger.warning(f"DeepSeek API недоступен для {cat_clean}: {e}")
        
        # 3. Локальная база
        if cat_clean in self.tariffs:
            t = self.tariffs[cat_clean]
            t.scheme = scheme
            return t
        for k, t in self.tariffs.items():
            if k in cat_clean or cat_clean in k:
                t.scheme = scheme
                return t
        t = self.tariffs['default']
        t.scheme = scheme
        return t
    
    def _parse_ym_tariffs(self, tariffs_data: List[Dict], category: str, scheme: str) -> Optional[Tariff]:
        if not tariffs_data:
            return None
        commission_rate = 0.12
        sorting_cost = 45.0
        delivery_rate = 0.045
        acquiring_rate = 0.016
        
        for t in tariffs_data:
            t_type = t.get("type", "")
            amount = t.get("amount", 0)
            params = {p.get("name", "").lower(): p.get("value", "") for p in t.get("parameters", [])}
            
            if t_type == "FEE":
                if params.get("valuetype", "") == "relative":
                    commission_rate = amount / 100.0
            elif t_type == "SORTING":
                sorting_cost = amount
            elif t_type == "DELIVERY_TO_CUSTOMER":
                if params.get("valuetype", "") == "relative":
                    delivery_rate = amount / 100.0
            elif t_type == "PAYMENT_TRANSFER":
                if params.get("valuetype", "") == "relative":
                    acquiring_rate = amount / 100.0
        
        return Tariff(
            category=category,
            commission_rate=commission_rate,
            min_commission=0,
            sorting_cost=sorting_cost,
            delivery_rate=delivery_rate,
            delivery_min=60,
            delivery_max=500,
            acquiring_transfer_rate=acquiring_rate,
            acquiring_sku_cost=0.12,
            source=f"API Яндекс Маркета ({scheme})",
            scheme=scheme
        )
    
    def to_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame([{
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
        } for k, t in self.tariffs.items()])

# ============================================================================
# БЛОК 3: ВЕКТОРИЗОВАННЫЙ ФИНАНСОВЫЙ ДВИЖОК (50+ МЕТРИК)
# ============================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def run_calculations_cached(df_hash: str, tax_label: str, tariffs_snapshot: str,
                            scheme_label: str, payment_frequency: str) -> pd.DataFrame:
    df = st.session_state.main_df.copy()
    if df.empty:
        return df
    
    tax_system = TaxSystem.by_label(tax_label)
    scheme = scheme_label.split(" ")[0]
    
    payment_rates = {
        "Ежемесячно (1.0%)": 0.01,
        "Раз в 2 недели (1.3%)": 0.013,
        "Еженедельно, 4 недели отсрочка (1.6%)": 0.016,
        "Еженедельно, 2 недели отсрочка (2.3%)": 0.023,
        "Еженедельно, 1 неделя отсрочка (2.8%)": 0.028,
        "Ежедневно (3.3%)": 0.033,
    }
    payment_transfer_rate = payment_rates.get(payment_frequency, 0.016)
    
    manager = HybridTariffManager()
    
    if 'artikul' in df.columns:
        df['artikul'] = df['artikul'].astype(str).apply(fix_double_utf8)
    if 'category' in df.columns:
        df['category'] = df['category'].astype(str).apply(fix_double_utf8)
    
    required_cols = {
        'selling_price': 0.0, 'cogs': 0.0, 'weight_kg': 0.0,
        'length_cm': 0.0, 'width_cm': 0.0, 'height_cm': 0.0,
        'packaging_cost': 0.0, 'marketing_budget_per_unit': 0.0,
        'daily_sales': 0.0, 'stock_depth_days': 0.0,
        'first_mile_cost': 0.0, 'commission': 0.0,
        'return_cost': 0.0, 'warehouse_cost': 0.0,
        'volume_liters': 0.0, 'quantity_per_order': 1.0
    }
    for col, default in required_cols.items():
        if col not in df.columns:
            df[col] = default
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(default)
    
    comm_rates, min_comms, sorting_costs = [], [], []
    delivery_rates, delivery_mins, delivery_maxs = [], [], []
    acq_transfer_rates, acq_sku_costs = [], []
    ret_rates, ret_processing = [], []
    storage_fees, special_rates = [], []
    middle_mile_bases, middle_mile_kgs = [], []
    
    for cat in df.get('category', ['default'] * len(df)):
        t = manager.get_best_tariff(cat, scheme)
        comm_rates.append(t.commission_rate)
        min_comms.append(t.min_commission)
        sorting_costs.append(t.sorting_cost)
        delivery_rates.append(t.delivery_rate)
        delivery_mins.append(t.delivery_min)
        delivery_maxs.append(t.delivery_max)
        acq_transfer_rates.append(payment_transfer_rate)
        acq_sku_costs.append(t.acquiring_sku_cost)
        ret_rates.append(t.return_rate)
        ret_processing.append(t.return_processing)
        storage_fees.append(t.storage_fee_per_day)
        special_rates.append(t.special_tariff_rate)
        middle_mile_bases.append(t.middle_mile_base)
        middle_mile_kgs.append(t.middle_mile_per_kg)
    
    comm_rates = np.array(comm_rates)
    delivery_rates = np.array(delivery_rates)
    acq_transfer_rates = np.array(acq_transfer_rates)
    ret_rates = np.array(ret_rates)
    special_rates = np.array(special_rates)
    
    # === РАСЧЁТЫ ПО ФОРМУЛАМ ЯНДЕКС МАРКЕТА ===
    vol_weight = (df['length_cm'] * df['width_cm'] * df['height_cm']) / 5000.0
    df['billable_weight'] = np.maximum(df['weight_kg'], vol_weight)
    df['billable_weight'] = np.ceil(df['billable_weight'] * 2) / 2
    
    df['is_special_tariff'] = (df['selling_price'] <= 300) & (df['volume_liters'] <= 5)
    
    # 1. КОМИССИЯ
    df['commission'] = np.where(
        df['is_special_tariff'],
        df['selling_price'] * special_rates,
        np.where(df['commission'] == 0, np.maximum(df['selling_price'] * comm_rates, min_comms), df['commission'])
    )
    
    # 2. ДОСТАВКА ПОКУПАТЕЛЮ
    df['delivery_to_customer'] = np.where(
        df['is_special_tariff'], 0.0,
        np.clip(df['selling_price'] * delivery_rates, np.array(delivery_mins), np.array(delivery_maxs))
    )
    
    # 3. СРЕДНЯЯ МИЛЯ
    df['middle_mile_cost'] = np.where(
        df['is_special_tariff'], 0.0,
        np.array([manager.get_middle_mile_cost(w, scheme) for w in df['billable_weight']])
    )
    
    # 4. ОБРАБОТКА
    df['sorting_cost'] = np.where(df['is_special_tariff'], 0.0, np.where(scheme == 'FBS', np.array(sorting_costs), 0.0))
    
    # 5. ЭКВАЙРИНГ
    df['acquiring_sku_cost'] = np.where(df['quantity_per_order'] > 0, np.array(acq_sku_costs) / df['quantity_per_order'], np.array(acq_sku_costs))
    df['acquiring_transfer_cost'] = df['selling_price'] * acq_transfer_rates
    df['acquiring_cost'] = df['acquiring_sku_cost'] + df['acquiring_transfer_cost']
    
    # 6. ВОЗВРАТЫ
    df['return_processing_cost'] = np.where(df['is_special_tariff'], 0.0, np.array(ret_processing))
    df['return_delivery_cost'] = np.where(df['is_special_tariff'], 0.0, df['middle_mile_cost'] * ret_rates)
    df['return_cost'] = np.where(df['return_cost'] == 0, df['return_processing_cost'] + df['return_delivery_cost'], df['return_cost'])
    
    # 7. УПАКОВКА
    df['pick_pack_cost'] = 35.0
    
    # 8. ХРАНЕНИЕ
    df['warehouse_cost'] = np.where(df['warehouse_cost'] == 0, (df['stock_depth_days'] * df['daily_sales']) * np.array(storage_fees), df['warehouse_cost'])
    
    # 9. ПЕРВАЯ МИЛЯ
    df['first_mile_cost'] = np.where(df['first_mile_cost'] == 0, np.array(middle_mile_bases) + (df['billable_weight'] * np.array(middle_mile_kgs)), df['first_mile_cost'])
    
    # === ИТОГОВЫЕ РАСЧЁТЫ ===
    df['fixed_operational_costs'] = (
        df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + 
        df['packaging_cost'] + df['return_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    )
    df['marketplace_fees'] = df['commission'] + df['delivery_to_customer'] + df['middle_mile_cost'] + df['sorting_cost'] + df['acquiring_cost']
    df['pre_tax_expenses'] = df['fixed_operational_costs'] + df['marketplace_fees']
    
    if tax_system.base == "revenue":
        df['tax_cost'] = df['selling_price'] * tax_system.rate
    elif tax_system.base == "profit":
        pre_tax_profit = df['selling_price'] - df['pre_tax_expenses']
        df['tax_cost'] = np.maximum(pre_tax_profit, 0) * tax_system.rate
    elif tax_system.base == "profit_vat":
        vat = df['selling_price'] * 0.20 / 1.20
        pre_tax_profit = (df['selling_price'] - vat) - df['pre_tax_expenses']
        df['tax_cost'] = np.maximum(pre_tax_profit, 0) * tax_system.rate
    
    df['total_expenses'] = df['pre_tax_expenses'] + df['tax_cost']
    df['gross_profit'] = df['selling_price'] - df['total_expenses']
    df['margin_percent'] = np.where(df['selling_price'] > 0, (df['gross_profit'] / df['selling_price']) * 100, 0.0)
    df['operating_profit'] = df['selling_price'] - df['pre_tax_expenses']
    df['operating_margin'] = np.where(df['selling_price'] > 0, (df['operating_profit'] / df['selling_price']) * 100, 0.0)
    
    # === РЕКОМЕНДОВАННЫЕ ЦЕНЫ ===
    tax_factor = tax_system.rate if tax_system.base == "revenue" else 0.0
    variable_fees_special = special_rates + acq_transfer_rates + tax_factor
    variable_fees_normal = comm_rates + delivery_rates + acq_transfer_rates + tax_factor
    df['variable_fees_rate'] = np.where(df['is_special_tariff'], variable_fees_special, variable_fees_normal)
    denom = np.where((1.0 - df['variable_fees_rate']) <= 0.01, 0.5, 1.0 - df['variable_fees_rate'])
    
    fixed_no_return = df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + df['packaging_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    df['rec_price_min'] = fixed_no_return / denom
    df['rec_price_10'] = fixed_no_return / (denom - 0.10)
    df['rec_price_15'] = fixed_no_return / (denom - 0.15)
    df['rec_price_20'] = fixed_no_return / (denom - 0.20)
    df['rec_price_25'] = fixed_no_return / (denom - 0.25)
    df['rec_price_30'] = fixed_no_return / (denom - 0.30)
    
    # === ДОП. МЕТРИКИ ===
    df['variable_costs'] = df['commission'] + df['delivery_to_customer'] + df['middle_mile_cost'] + df['sorting_cost'] + df['acquiring_cost'] + df['return_cost']
    df['fixed_costs'] = df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + df['packaging_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    df['contribution_margin'] = df['selling_price'] - df['variable_costs']
    df['contribution_margin_percent'] = np.where(df['selling_price'] > 0, (df['contribution_margin'] / df['selling_price']) * 100, 0.0)
    df['gross_margin_before_tax'] = df['selling_price'] - df['pre_tax_expenses']
    df['net_margin_after_tax'] = df['margin_percent']
    
    for prefix, numerator in [
        ('cogs', 'cogs'), ('commission', 'commission'), ('delivery', 'delivery_to_customer'),
        ('middle_mile', 'middle_mile_cost'), ('sorting', 'sorting_cost'), ('acquiring', 'acquiring_cost'),
        ('marketing', 'marketing_budget_per_unit'), ('return', 'return_cost'), ('tax', 'tax_cost')
    ]:
        df[f'{prefix}_percent'] = np.where(df['selling_price'] > 0, (df[numerator] / df['selling_price']) * 100, 0.0)
    
    df['logistics_percent'] = np.where(df['selling_price'] > 0, ((df['first_mile_cost'] + df['delivery_to_customer'] + df['middle_mile_cost']) / df['selling_price']) * 100, 0.0)
    df['total_fees_percent'] = df['cogs_percent'] + df['logistics_percent'] + df['commission_percent'] + df['marketing_percent'] + df['return_percent'] + df['tax_percent']
    df['roi_percent'] = np.where(df['cogs'] > 0, ((df['gross_profit'] / df['cogs']) * 100), 0.0)
    df['markup_percent'] = np.where(df['cogs'] > 0, ((df['selling_price'] - df['cogs']) / df['cogs']) * 100, 0.0)
    df['break_even_units'] = np.where(df['contribution_margin'] > 0, df['fixed_costs'] / df['contribution_margin'], 0.0)
    df['safety_margin_percent'] = np.where(df['selling_price'] > df['rec_price_min'], ((df['selling_price'] - df['rec_price_min']) / df['selling_price']) * 100, 0.0)
    df['cost_per_kg'] = np.where(df['billable_weight'] > 0, (df['first_mile_cost'] + df['delivery_to_customer'] + df['middle_mile_cost']) / df['billable_weight'], 0.0)
    df['revenue_per_kg'] = np.where(df['billable_weight'] > 0, df['selling_price'] / df['billable_weight'], 0.0)
    df['profit_per_kg'] = np.where(df['billable_weight'] > 0, df['gross_profit'] / df['billable_weight'], 0.0)
    df['efficiency_score'] = df['margin_percent'] * 0.4 + df['roi_percent'] * 0.3 + df['safety_margin_percent'] * 0.3
    df['abc_category'] = np.where(df['daily_sales'] >= 10, 'A', np.where(df['daily_sales'] >= 3, 'B', 'C'))
    df['xyz_category'] = np.where(df['margin_percent'] >= 20, 'X', np.where(df['margin_percent'] >= 10, 'Y', 'Z'))
    df['abc_xyz'] = df['abc_category'] + df['xyz_category']
    df['profitability_status'] = np.where(
        df['gross_profit'] > 0,
        np.where(df['margin_percent'] >= 20, 'Высокомаржинальный',
                 np.where(df['margin_percent'] >= 10, 'Среднемаржинальный', 'Низкомаржинальный')),
        np.where(df['gross_profit'] == 0, 'На грани', 'Убыточный')
    )
    
    money_columns = [
        'commission', 'delivery_to_customer', 'middle_mile_cost', 'sorting_cost',
        'first_mile_cost', 'acquiring_cost', 'acquiring_sku_cost', 'acquiring_transfer_cost',
        'return_cost', 'return_processing_cost', 'return_delivery_cost',
        'pick_pack_cost', 'warehouse_cost', 'fixed_operational_costs',
        'marketplace_fees', 'pre_tax_expenses', 'tax_cost', 'total_expenses',
        'gross_profit', 'operating_profit',
        'rec_price_min', 'rec_price_10', 'rec_price_15', 'rec_price_20', 'rec_price_25', 'rec_price_30',
        'variable_costs', 'fixed_costs', 'contribution_margin',
        'cost_per_kg', 'revenue_per_kg', 'profit_per_kg'
    ]
    for col in money_columns:
        if col in df.columns:
            df[col] = df[col].apply(money_round)
    
    percent_columns = [
        'margin_percent', 'operating_margin', 'contribution_margin_percent',
        'cogs_percent', 'logistics_percent', 'commission_percent', 'delivery_percent',
        'middle_mile_percent', 'sorting_percent', 'acquiring_percent',
        'marketing_percent', 'return_percent', 'tax_percent', 'total_fees_percent',
        'roi_percent', 'markup_percent', 'safety_margin_percent', 'efficiency_score',
        'variable_fees_rate'
    ]
    for col in percent_columns:
        if col in df.columns:
            df[col] = df[col].apply(percent_round)
    
    return df

# ============================================================================
# БЛОК 4: ЭКСПОРТЁР С ЖИВЫМИ ФОРМУЛАМИ
# ============================================================================
class UltimateExcelExporter:
    @staticmethod
    def _get_col_letter(idx: int) -> str:
        res = ""
        while idx >= 0:
            res = chr(idx % 26 + 65) + res
            idx = idx // 26 - 1
        return res
    
    @staticmethod
    def export_max_info(df: pd.DataFrame, tax_label: str, scheme_label: str) -> bytes:
        if not OPENPYXL_AVAILABLE or df.empty:
            return b""
        
        wb = Workbook()
        
        # === 1. СВОДКА ===
        ws_summary = wb.active
        ws_summary.title = "Сводный Дашборд"
        ws_summary.merge_cells('A1:E1')
        cell = ws_summary.cell(1, 1, "СВОДНЫЙ ФИНАНСОВЫЙ ОТЧЁТ — ЯНДЕКС МАРКЕТ")
        cell.font = Font(size=18, bold=True, color="1F4E78")
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
        ws_summary.cell(2, 1, f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M')}")
        ws_summary.cell(3, 1, f"Схема: {scheme_label}")
        ws_summary.cell(3, 3, f"Налог: {tax_label}")
        
        metrics = [
            ("Всего SKU", len(df)),
            ("Общая выручка", df['selling_price'].sum()),
            ("Общие расходы", df['total_expenses'].sum()),
            ("ОБЩАЯ ПРИБЫЛЬ", df['gross_profit'].sum()),
            ("Средняя маржинальность", df['margin_percent'].mean()),
            ("Средний ROI", df['roi_percent'].mean()),
            ("Прибыльных SKU", (df['gross_profit'] > 0).sum()),
            ("Убыточных SKU", (df['gross_profit'] < 0).sum()),
        ]
        
        row = 5
        for label, value in metrics:
            ws_summary.cell(row, 1, label).font = Font(bold=True, size=11)
            val_cell = ws_summary.cell(row, 2, value)
            val_cell.font = Font(bold=True, size=11)
            if isinstance(value, float) and ('маржинальность' in label or 'ROI' in label):
                val_cell.number_format = '0.00"%"'
            elif isinstance(value, (int, float)):
                val_cell.number_format = '#,##0.00 "₽"'
            row += 1
        
        ws_summary.cell(row + 1, 1, "СТРУКТУРА РАСХОДОВ").font = Font(bold=True, size=12, color="1F4E78")
        expense_items = [
            ("Себестоимость", df['cogs'].sum()),
            ("Комиссия размещение", df['commission'].sum()),
            ("Доставка покупателю", df['delivery_to_customer'].sum()),
            ("Средняя миля", df['middle_mile_cost'].sum()),
            ("Обработка заказа", df['sorting_cost'].sum()),
            ("Эквайринг", df['acquiring_cost'].sum()),
            ("Возвраты", df['return_cost'].sum()),
            ("Упаковка/пикпак", df['pick_pack_cost'].sum()),
            ("Хранение", df['warehouse_cost'].sum()),
            ("Маркетинг", df['marketing_budget_per_unit'].sum()),
            ("Налоги", df['tax_cost'].sum()),
        ]
        row += 2
        for label, value in expense_items:
            ws_summary.cell(row, 1, label).font = Font(size=10)
            val_cell = ws_summary.cell(row, 2, value)
            val_cell.number_format = '#,##0.00 "₽"'
            row += 1
        
        for col in ['A', 'B']:
            ws_summary.column_dimensions[col].width = 30
        
        # === 2. ДЕТАЛЬНЫЙ РАСЧЁТ С ЖИВЫМИ ФОРМУЛАМИ ===
        ws_detail = wb.create_sheet("Детальный расчет")
        
        priority_cols = [
            'artikul', 'category', 'selling_price', 'cogs', 'weight_kg', 'billable_weight',
            'volume_liters', 'is_special_tariff',
            'commission', 'delivery_to_customer', 'middle_mile_cost', 'sorting_cost',
            'acquiring_cost', 'acquiring_sku_cost', 'acquiring_transfer_cost',
            'return_cost', 'return_processing_cost', 'return_delivery_cost',
            'pick_pack_cost', 'warehouse_cost', 'first_mile_cost', 'packaging_cost',
            'marketing_budget_per_unit',
            'fixed_operational_costs', 'marketplace_fees', 'pre_tax_expenses',
            'tax_cost', 'total_expenses', 'gross_profit', 'operating_profit',
            'margin_percent', 'operating_margin', 'contribution_margin', 'contribution_margin_percent',
            'rec_price_min', 'rec_price_10', 'rec_price_15', 'rec_price_20', 'rec_price_25', 'rec_price_30',
            'cogs_percent', 'commission_percent', 'delivery_percent', 'middle_mile_percent',
            'sorting_percent', 'acquiring_percent', 'marketing_percent', 'return_percent', 'tax_percent',
            'logistics_percent', 'total_fees_percent',
            'roi_percent', 'markup_percent', 'break_even_units', 'safety_margin_percent',
            'efficiency_score', 'abc_xyz', 'profitability_status',
            'daily_sales', 'stock_depth_days', 'cost_per_kg', 'revenue_per_kg', 'profit_per_kg'
        ]
        
        cols = [c for c in priority_cols if c in df.columns]
        cols += [c for c in df.columns if c not in priority_cols]
        
        def get_letter(idx: int) -> str:
            return UltimateExcelExporter._get_col_letter(idx)
        
        # Заголовки
        for col_idx, col_name in enumerate(cols, 1):
            cell = ws_detail.cell(1, col_idx, col_name)
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = PatternFill(start_color="2E75B6", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        
        ws_detail.auto_filter.ref = f"A1:{get_letter(len(cols)-1)}{len(df)+1}"
        ws_detail.freeze_panes = 'B2'
        
        idx_map = {col: i for i, col in enumerate(cols)}
        
        def cell_ref(col_name: str, row_num: int) -> str:
            if col_name in idx_map:
                return f"{get_letter(idx_map[col_name])}{row_num}"
            return ""
        
        for r_idx, row_data in enumerate(df.itertuples(index=False), 2):
            for c_idx, value in enumerate(row_data, 1):
                cell = ws_detail.cell(r_idx, c_idx, value)
                cell.border = Border(bottom=Side(style="thin", color="E0E0E0"))
                col_name = cols[c_idx - 1]
                if isinstance(value, (int, float)):
                    if 'percent' in col_name or col_name in ['efficiency_score', 'variable_fees_rate']:
                        cell.number_format = '0.00"%"'
                    elif col_name in ['daily_sales', 'stock_depth_days', 'break_even_units', 'quantity_per_order']:
                        cell.number_format = '0'
                    elif col_name != 'is_special_tariff':
                        cell.number_format = '#,##0.00'
                
                if col_name == 'gross_profit' and isinstance(value, (int, float)):
                    if value < 0:
                        cell.fill = PatternFill(start_color="FFC7CE", fill_type="solid")
                        cell.font = Font(color="9C0006", bold=True)
                    elif value > 0:
                        cell.fill = PatternFill(start_color="C6EFCE", fill_type="solid")
                        cell.font = Font(color="006100", bold=True)
            
            # ЖИВЫЕ ФОРМУЛЫ
            if all(c in idx_map for c in ['billable_weight', 'weight_kg', 'length_cm', 'width_cm', 'height_cm']):
                formula = f"=MAX({cell_ref('weight_kg', r_idx)},{cell_ref('length_cm', r_idx)}*{cell_ref('width_cm', r_idx)}*{cell_ref('height_cm', r_idx)}/5000)"
                ws_detail.cell(r_idx, idx_map['billable_weight'] + 1).value = formula
            
            if all(c in idx_map for c in ['is_special_tariff', 'selling_price', 'volume_liters']):
                formula = f"=IF(AND({cell_ref('selling_price', r_idx)}<=300,{cell_ref('volume_liters', r_idx)}<=5),TRUE,FALSE)"
                ws_detail.cell(r_idx, idx_map['is_special_tariff'] + 1).value = formula
            
            if all(c in idx_map for c in ['acquiring_cost', 'acquiring_sku_cost', 'acquiring_transfer_cost']):
                formula = f"={cell_ref('acquiring_sku_cost', r_idx)}+{cell_ref('acquiring_transfer_cost', r_idx)}"
                ws_detail.cell(r_idx, idx_map['acquiring_cost'] + 1).value = formula
            
            if all(c in idx_map for c in ['return_cost', 'return_processing_cost', 'return_delivery_cost']):
                formula = f"={cell_ref('return_processing_cost', r_idx)}+{cell_ref('return_delivery_cost', r_idx)}"
                ws_detail.cell(r_idx, idx_map['return_cost'] + 1).value = formula
            
            fixed_parts = ['cogs', 'first_mile_cost', 'pick_pack_cost', 'packaging_cost', 'return_cost', 'marketing_budget_per_unit', 'warehouse_cost']
            if all(c in idx_map for c in fixed_parts) and 'fixed_operational_costs' in idx_map:
                formula = f"={'+'.join([cell_ref(c, r_idx) for c in fixed_parts])}"
                ws_detail.cell(r_idx, idx_map['fixed_operational_costs'] + 1).value = formula
            
            fee_parts = ['commission', 'delivery_to_customer', 'middle_mile_cost', 'sorting_cost', 'acquiring_cost']
            if all(c in idx_map for c in fee_parts) and 'marketplace_fees' in idx_map:
                formula = f"={'+'.join([cell_ref(c, r_idx) for c in fee_parts])}"
                ws_detail.cell(r_idx, idx_map['marketplace_fees'] + 1).value = formula
            
            if all(c in idx_map for c in ['fixed_operational_costs', 'marketplace_fees', 'pre_tax_expenses']):
                formula = f"={cell_ref('fixed_operational_costs', r_idx)}+{cell_ref('marketplace_fees', r_idx)}"
                ws_detail.cell(r_idx, idx_map['pre_tax_expenses'] + 1).value = formula
            
            if all(c in idx_map for c in ['pre_tax_expenses', 'tax_cost', 'total_expenses']):
                formula = f"={cell_ref('pre_tax_expenses', r_idx)}+{cell_ref('tax_cost', r_idx)}"
                ws_detail.cell(r_idx, idx_map['total_expenses'] + 1).value = formula
            
            if all(c in idx_map for c in ['gross_profit', 'selling_price', 'total_expenses']):
                formula = f"={cell_ref('selling_price', r_idx)}-{cell_ref('total_expenses', r_idx)}"
                ws_detail.cell(r_idx, idx_map['gross_profit'] + 1).value = formula
            
            if all(c in idx_map for c in ['operating_profit', 'selling_price', 'pre_tax_expenses']):
                formula = f"={cell_ref('selling_price', r_idx)}-{cell_ref('pre_tax_expenses', r_idx)}"
                ws_detail.cell(r_idx, idx_map['operating_profit'] + 1).value = formula
            
            if all(c in idx_map for c in ['margin_percent', 'gross_profit', 'selling_price']):
                formula = f"=IF({cell_ref('selling_price', r_idx)}>0,{cell_ref('gross_profit', r_idx)}/{cell_ref('selling_price', r_idx)},0)"
                ws_detail.cell(r_idx, idx_map['margin_percent'] + 1).value = formula
                ws_detail.cell(r_idx, idx_map['margin_percent'] + 1).number_format = '0.00%'
            
            if all(c in idx_map for c in ['contribution_margin', 'selling_price', 'variable_costs']):
                formula = f"={cell_ref('selling_price', r_idx)}-{cell_ref('variable_costs', r_idx)}"
                ws_detail.cell(r_idx, idx_map['contribution_margin'] + 1).value = formula
            
            if all(c in idx_map for c in ['roi_percent', 'gross_profit', 'cogs']):
                formula = f"=IF({cell_ref('cogs', r_idx)}>0,{cell_ref('gross_profit', r_idx)}/{cell_ref('cogs', r_idx)},0)"
                ws_detail.cell(r_idx, idx_map['roi_percent'] + 1).value = formula
                ws_detail.cell(r_idx, idx_map['roi_percent'] + 1).number_format = '0.00%'
            
            if all(c in idx_map for c in ['markup_percent', 'selling_price', 'cogs']):
                formula = f"=IF({cell_ref('cogs', r_idx)}>0,({cell_ref('selling_price', r_idx)}-{cell_ref('cogs', r_idx)})/{cell_ref('cogs', r_idx)},0)"
                ws_detail.cell(r_idx, idx_map['markup_percent'] + 1).value = formula
                ws_detail.cell(r_idx, idx_map['markup_percent'] + 1).number_format = '0.00%'
            
            if all(c in idx_map for c in ['safety_margin_percent', 'selling_price', 'rec_price_min']):
                formula = f"=IF({cell_ref('selling_price', r_idx)}>{cell_ref('rec_price_min', r_idx)},({cell_ref('selling_price', r_idx)}-{cell_ref('rec_price_min', r_idx)})/{cell_ref('selling_price', r_idx)},0)"
                ws_detail.cell(r_idx, idx_map['safety_margin_percent'] + 1).value = formula
                ws_detail.cell(r_idx, idx_map['safety_margin_percent'] + 1).number_format = '0.00%'
            
            if all(c in idx_map for c in ['break_even_units', 'fixed_costs', 'contribution_margin']):
                formula = f"=IF({cell_ref('contribution_margin', r_idx)}>0,{cell_ref('fixed_costs', r_idx)}/{cell_ref('contribution_margin', r_idx)},0)"
                ws_detail.cell(r_idx, idx_map['break_even_units'] + 1).value = formula
            
            if all(c in idx_map for c in ['cost_per_kg', 'first_mile_cost', 'delivery_to_customer', 'middle_mile_cost', 'billable_weight']):
                formula = f"=IF({cell_ref('billable_weight', r_idx)}>0,({cell_ref('first_mile_cost', r_idx)}+{cell_ref('delivery_to_customer', r_idx)}+{cell_ref('middle_mile_cost', r_idx)})/{cell_ref('billable_weight', r_idx)},0)"
                ws_detail.cell(r_idx, idx_map['cost_per_kg'] + 1).value = formula
            
            if all(c in idx_map for c in ['revenue_per_kg', 'selling_price', 'billable_weight']):
                formula = f"=IF({cell_ref('billable_weight', r_idx)}>0,{cell_ref('selling_price', r_idx)}/{cell_ref('billable_weight', r_idx)},0)"
                ws_detail.cell(r_idx, idx_map['revenue_per_kg'] + 1).value = formula
            
            if all(c in idx_map for c in ['profit_per_kg', 'gross_profit', 'billable_weight']):
                formula = f"=IF({cell_ref('billable_weight', r_idx)}>0,{cell_ref('gross_profit', r_idx)}/{cell_ref('billable_weight', r_idx)},0)"
                ws_detail.cell(r_idx, idx_map['profit_per_kg'] + 1).value = formula
        
        # Условное форматирование
        if 'gross_profit' in idx_map:
            col_letter = get_letter(idx_map['gross_profit'])
            green_rule = FormulaRule(formula=[f'{col_letter}2>0'], fill=PatternFill(start_color="C6EFCE", fill_type="solid"), font=Font(color="006100", bold=True))
            red_rule = FormulaRule(formula=[f'{col_letter}2<0'], fill=PatternFill(start_color="FFC7CE", fill_type="solid"), font=Font(color="9C0006", bold=True))
            ws_detail.conditional_formatting.add(f'{col_letter}2:{col_letter}{len(df)+1}', green_rule)
            ws_detail.conditional_formatting.add(f'{col_letter}2:{col_letter}{len(df)+1}', red_rule)
        
        if 'margin_percent' in idx_map:
            col_letter = get_letter(idx_map['margin_percent'])
            ws_detail.conditional_formatting.add(f'{col_letter}2:{col_letter}{len(df)+1}', DataBarRule(start_type='min', end_type='max', color="638EC6", showValue=True))
        
        if 'efficiency_score' in idx_map:
            col_letter = get_letter(idx_map['efficiency_score'])
            ws_detail.conditional_formatting.add(f'{col_letter}2:{col_letter}{len(df)+1}', ColorScaleRule(start_type='min', start_color='F8696B', mid_type='percentile', mid_value=50, mid_color='FFEB84', end_type='max', end_color='63BE7B'))
        
        for col_idx, col_name in enumerate(cols, 1):
            width = 18
            if col_name in ['artikul', 'category', 'profitability_status', 'abc_xyz']:
                width = 22
            elif 'percent' in col_name or col_name in ['efficiency_score']:
                width = 14
            ws_detail.column_dimensions[get_letter(col_idx - 1)].width = width
        
        # === 3. РЕКОМЕНДОВАННЫЕ ЦЕНЫ ===
        ws_prices = wb.create_sheet("Рекомендованные цены")
        price_cols = ['artikul', 'category', 'selling_price', 'cogs', 'gross_profit', 'margin_percent', 'rec_price_min', 'rec_price_10', 'rec_price_15', 'rec_price_20', 'rec_price_25', 'rec_price_30', 'profitability_status']
        price_cols = [c for c in price_cols if c in df.columns]
        for col_idx, col_name in enumerate(price_cols, 1):
            cell = ws_prices.cell(1, col_idx, col_name)
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = PatternFill(start_color="1F4E78", fill_type="solid")
        for r_idx, row_data in enumerate(df[price_cols].itertuples(index=False), 2):
            for c_idx, value in enumerate(row_data, 1):
                cell = ws_prices.cell(r_idx, c_idx, value)
                col_name = price_cols[c_idx - 1]
                if 'percent' in col_name:
                    cell.number_format = '0.00%'
                elif col_name in ['selling_price', 'cogs', 'gross_profit'] or col_name.startswith('rec_price'):
                    cell.number_format = '#,##0.00'
                if col_name == 'gross_profit' and isinstance(value, (int, float)):
                    cell.fill = PatternFill(start_color="C6EFCE" if value > 0 else "FFC7CE", fill_type="solid")
                    cell.font = Font(color="006100" if value > 0 else "9C0006", bold=True)
        for col_idx in range(1, len(price_cols) + 1):
            ws_prices.column_dimensions[get_letter(col_idx - 1)].width = 18
        
        # === 4. ABC-XYZ ===
        ws_abc = wb.create_sheet("ABC-XYZ Анализ")
        abc_data = df.groupby('abc_xyz').agg({'artikul': 'count', 'selling_price': 'sum', 'gross_profit': 'sum', 'margin_percent': 'mean', 'daily_sales': 'sum'}).reset_index()
        abc_data.columns = ['ABC-XYZ', 'Кол-во SKU', 'Выручка', 'Прибыль', 'Ср. маржа %', 'Продажи/день']
        for col_idx, col_name in enumerate(abc_data.columns, 1):
            cell = ws_abc.cell(1, col_idx, col_name)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="2E75B6", fill_type="solid")
        for r_idx, row_data in enumerate(abc_data.itertuples(index=False), 2):
            for c_idx, value in enumerate(row_data, 1):
                cell = ws_abc.cell(r_idx, c_idx, value)
                if c_idx > 2:
                    cell.number_format = '#,##0.00'
        for col_idx in range(1, len(abc_data.columns) + 1):
            ws_abc.column_dimensions[get_letter(col_idx - 1)].width = 18
        
        # === 5. СПРАВКА ===
        ws_help = wb.create_sheet("Справка по тарифам")
        help_text = [
            ["ТАРИФЫ ЯНДЕКС МАРКЕТА (август 2026)", ""],
            ["", ""],
            ["1. КОМИССИЯ ЗА РАЗМЕЩЕНИЕ", "Зависит от категории и модели (FBS/FBY/DBS/Экспресс)"],
            ["   Автозапчасти", "5-19%"], ["   Электроника", "10-23%"], ["   Одежда/Обувь", "14.5-23%"], ["   Детские товары", "5-20%"],
            ["", ""],
            ["2. ДОСТАВКА ПОКУПАТЕЛЮ", "4.5% от цены, мин. 60₽, макс. 500₽"],
            ["   Экспресс (до 25кг, <200см)", "6% от цены, мин. 80₽, макс. 500₽"], ["   Экспресс (свыше)", "700₽ за товар"],
            ["", ""],
            ["3. СРЕДНЯЯ МИЛЯ (FBS/FBY)", "Зависит от веса/объёмного веса"],
            ["   FBS: до 4кг", "80-190₽"], ["   FBS: 4-10кг", "270-440₽"], ["   FBS: 10-20кг", "530-800₽"], ["   FBS: свыше 20кг", "1000-3500₽"],
            ["   FBY: до 4кг", "60-160₽"], ["   FBY: 4-10кг", "230-370₽"], ["   FBY: 10-20кг", "440-670₽"], ["   FBY: свыше 20кг", "840-3500₽"],
            ["", ""],
            ["4. ОБРАБОТКА ЗАКАЗА (FBS)", "45₽ за заказ"],
            ["", ""],
            ["5. ЭКВАЙРИНГ", ""],
            ["   Приём платежа", "0.12₽ за SKU в заказе"], ["   Перевод денег", "1.0-3.3% в зависимости от частоты выплат"],
            ["", ""],
            ["6. ВОЗВРАТЫ", ""],
            ["   Обработка в СЦ", "15₽ за отправление"], ["   Доставка обратно", "По тарифу средней мили"],
            ["", ""],
            ["7. СПЕЦТАРИФ (≤300₽ и ≤5л)", "42% от цены (FBS), 35% (FBY)"],
            ["   Включает: размещение, доставку, среднюю милю, возврат", ""],
            ["", ""],
            ["8. НАЛОГИ (2026)", ""],
            ["   УСН 6%", "6% от выручки"], ["   УСН 15%", "15% от (выручка - расходы)"],
            ["   АУСН 8%", "8% от выручки"], ["   АУСН 20%", "20% от (выручка - расходы)"],
            ["   ОСН", "НДС 20% + налог на прибыль"],
            ["", ""],
            ["9. СКИДКА ЗА БЫСТРУЮ ОТГРУЗКУ FBS", ""],
            ["   <=36 часов", "-4 п.п. от ставки размещения"], ["   <=28 часов", "-7 п.п. от ставки размещения"],
        ]
        for r_idx, (col1, col2) in enumerate(help_text, 1):
            ws_help.cell(r_idx, 1, col1).font = Font(bold=(r_idx == 1), size=11 if r_idx == 1 else 10)
            ws_help.cell(r_idx, 2, col2).font = Font(size=10)
            if r_idx == 1:
                ws_help.cell(r_idx, 1).fill = PatternFill(start_color="1F4E78", fill_type="solid")
                ws_help.cell(r_idx, 1).font = Font(bold=True, color="FFFFFF", size=12)
        ws_help.column_dimensions['A'].width = 45
        ws_help.column_dimensions['B'].width = 60
        
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.getvalue()

# ============================================================================
# БЛОК 5: НОРМАЛИЗАТОР ДАННЫХ
# ============================================================================
class UniversalDataNormalizer:
    COLUMN_MAPPING_DICTIONARY = {
        'artikul': ['artikul', 'артикул', 'код товара', 'sku', 'offer_id', 'id', 'товар', 'sku_id', 'vendor_code'],
        'category': ['category', 'категория', 'группа', 'тип товара', 'предмет', 'category_name', 'категория товара'],
        'selling_price': ['selling_price', 'цена продажи', 'цена', 'price', 'реализация', 'выручка', 'цена, руб', 'цена руб'],
        'cogs': ['cogs', 'себестоимость', 'закупка', 'cost', 'себестоимость р.', 'purchase_price', 'закупочная цена'],
        'daily_sales': ['daily_sales', 'заказы, шт.', 'продажи, шт.', 'quantity', 'sales_count', 'заказы', 'sales_per_day', 'продажи в день'],
        'weight_kg': ['weight_kg', 'вес', 'weight', 'вес кг', 'вес, кг', 'mass'],
        'length_cm': ['length_cm', 'длина', 'length', 'длина, см', 'длина см'],
        'width_cm': ['width_cm', 'ширина', 'width', 'ширина, см', 'ширина см'],
                'height_cm': ['height_cm', 'высота', 'height', 'высота, см', 'высота см'],
        'volume_liters': ['volume_liters', 'объем', 'volume', 'объем, л', 'объем л', 'литры'],
        'packaging_cost': ['packaging_cost', 'упаковка', 'packaging', 'стоимость упаковки'],
        'first_mile_cost': ['first_mile_cost', 'магистраль', 'логистика', 'доставка', 'первая миля'],
        'commission': ['commission', 'комиссия', 'marketplace_fee', 'комиссия маркетплейса'],
        'return_cost': ['return_cost', 'возвраты', 'возвраты ₽', 'return_fee'],
        'marketing_budget_per_unit': ['marketing_budget_per_unit', 'рекламные затраты', 'реклама', 'marketing', 'дрр', 'drr'],
        'warehouse_cost': ['warehouse_cost', 'стоимость хранения', 'хранение', 'storage', 'storage_cost'],
        'stock_depth_days': ['stock_depth_days', 'глубина запаса', 'stock_days', 'дни запаса'],
        'quantity_per_order': ['quantity_per_order', 'количество в заказе', 'квант', 'quantum'],
    }
    
    @classmethod
    def normalize_dataframe(cls, raw_df: pd.DataFrame) -> pd.DataFrame:
        if raw_df.empty:
            return raw_df
        df = raw_df.copy()
        df.columns = [str(col).strip().lower() for col in df.columns]
        final_data = {}
        
        for target_col, synonyms in cls.COLUMN_MAPPING_DICTIONARY.items():
            found = False
            for synonym in synonyms:
                if synonym in df.columns:
                    final_data[target_col] = df[synonym]
                    found = True
                    break
            if not found:
                if target_col in ['artikul', 'category']:
                    final_data[target_col] = "Кастомный SKU" if target_col == 'artikul' else "автозапчасти"
                else:
                    final_data[target_col] = 0.0
        
        normalized_df = pd.DataFrame(final_data)
        
        numeric_cols = ['selling_price', 'cogs', 'daily_sales', 'weight_kg', 'length_cm', 
                        'width_cm', 'height_cm', 'volume_liters', 'packaging_cost',
                        'first_mile_cost', 'commission', 'return_cost', 
                        'marketing_budget_per_unit', 'warehouse_cost', 'stock_depth_days',
                        'quantity_per_order']
        
        for col in numeric_cols:
            if col in normalized_df.columns:
                normalized_df[col] = (normalized_df[col]
                    .astype(str)
                    .str.replace(r'\s+', '', regex=True)
                    .str.replace(',', '.')
                    .str.replace('%', '', regex=False))
                normalized_df[col] = pd.to_numeric(normalized_df[col], errors='coerce').fillna(0.0)
                if col in ['cogs', 'first_mile_cost', 'commission', 'return_cost', 
                           'marketing_budget_per_unit', 'warehouse_cost', 'packaging_cost']:
                    normalized_df[col] = normalized_df[col].abs()
        
        normalized_df['artikul'] = normalized_df['artikul'].astype(str).str.strip()
        normalized_df['category'] = normalized_df['category'].astype(str).str.strip().str.lower()
        
        return normalized_df
    
    @classmethod
    def load_file_dynamically(cls, file_buffer: io.BytesIO, file_name: str) -> pd.DataFrame:
        try:
            if file_name.endswith('.csv'):
                return pd.read_csv(file_buffer, sep=None, engine='python', encoding='utf-8')
            elif file_name.endswith(('.xls', '.xlsx')):
                return pd.read_excel(file_buffer)
            elif file_name.endswith('.json'):
                return pd.read_json(file_buffer)
            else:
                raise ValueError("Неподдерживаемый формат расширения файла.")
        except UnicodeDecodeError:
            file_buffer.seek(0)
            return pd.read_csv(file_buffer, sep=None, engine='python', encoding='cp1251')

# ============================================================================
# БЛОК 6: ПРОФЕССИОНАЛЬНЫЕ ВИЗУАЛИЗАЦИИ
# ============================================================================
def render_waterfall_chart(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        return go.Figure()
    avg_row = df.mean(numeric_only=True)
    labels = ["Цена продажи", "Себестоимость", "Логистика (1-я + ср. + посл. миля)", 
              "Комиссия МП", "Обработка/упаковка", "Эквайринг", "Возвраты", "Маркетинг", 
              "Хранение", "Налог", "Чистая прибыль"]
    
    price = avg_row.get('selling_price', 0)
    cogs = -avg_row.get('cogs', 0)
    logistics = -(avg_row.get('first_mile_cost', 0) + avg_row.get('middle_mile_cost', 0) + avg_row.get('delivery_to_customer', 0))
    commission = -avg_row.get('commission', 0)
    sorting_pack = -(avg_row.get('sorting_cost', 0) + avg_row.get('pick_pack_cost', 0) + avg_row.get('packaging_cost', 0))
    acquiring = -avg_row.get('acquiring_cost', 0)
    returns = -avg_row.get('return_cost', 0)
    marketing = -avg_row.get('marketing_budget_per_unit', 0)
    storage = -avg_row.get('warehouse_cost', 0)
    tax = -avg_row.get('tax_cost', 0)
    profit = avg_row.get('gross_profit', 0)
    
    values = [price, cogs, logistics, commission, sorting_pack, acquiring, returns, marketing, storage, tax, profit]
    measure = ["absolute"] + ["relative"] * 9 + ["total"]
    
    fig = go.Figure(go.Waterfall(
        name="Unit Economics",
        orientation="v",
        measure=measure,
        x=labels,
        textposition="outside",
        text=[f"{v:,.0f} ₽" for v in values],
        y=values,
        connector={"line": {"color": "rgb(63, 63, 63)"}},
        increasing={"marker": {"color": "#2ca02c"}},
        decreasing={"marker": {"color": "#d62728"}},
        totals={"marker": {"color": "#1f77b4"}}
    ))
    fig.update_layout(
        title="Структура цены среднего товара (Waterfall)",
        showlegend=False,
        height=500
    )
    return fig

def render_margin_scatter(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        return go.Figure()
    fig = px.scatter(
        df, x="daily_sales", y="margin_percent", size="gross_profit", color="abc_xyz",
        hover_data=["artikul", "selling_price", "gross_profit", "profitability_status"],
        title="Матрица эффективности: Маржинальность (%) vs Продажи (шт/день)",
        labels={"daily_sales": "Продажи (шт/день)", "margin_percent": "Маржинальность (%)"},
        color_discrete_map={
            'AX': '#006100', 'AY': '#2ca02c', 'BX': '#1f77b4', 'BY': '#17becf',
            'CX': '#ff7f0e', 'CY': '#bcbd22', 'CZ': '#d62728', 'AZ': '#e377c2'
        }
    )
    fig.add_hline(y=0, line_dash="dash", line_color="red", annotation_text="Точка безубыточности")
    fig.add_vline(x=3, line_dash="dot", line_color="gray", annotation_text="B/C граница")
    fig.add_vline(x=10, line_dash="dot", line_color="gray", annotation_text="A/B граница")
    fig.update_layout(height=550)
    return fig

def render_pareto_chart(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        return go.Figure()
    df_sorted = df.sort_values('gross_profit', ascending=False).reset_index(drop=True)
    df_sorted['cumulative_profit_pct'] = df_sorted['gross_profit'].cumsum() / df_sorted['gross_profit'].sum() * 100
    
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=df_sorted.index,
        y=df_sorted['gross_profit'],
        name='Прибыль по SKU',
        marker_color=np.where(df_sorted['gross_profit'] >= 0, '#2ca02c', '#d62728')
    ))
    fig.add_trace(go.Scatter(
        x=df_sorted.index,
        y=df_sorted['cumulative_profit_pct'],
        name='Накопленная доля %',
        yaxis='y2',
        line=dict(color='#1f77b4', width=3)
    ))
    fig.update_layout(
        title="Кривая Парето: прибыльность SKU",
        xaxis_title="SKU (ранжировано по прибыли)",
        yaxis_title="Прибыль, ₽",
        yaxis2=dict(title="Накопленная доля %", overlaying='y', side='right', range=[0, 110]),
        height=500,
        showlegend=True
    )
    return fig

def render_efficiency_heatmap(df: pd.DataFrame) -> go.Figure:
    if df.empty:
        return go.Figure()
    pivot = df.pivot_table(values='margin_percent', index='abc_category', 
                            columns='xyz_category', aggfunc='mean')
    fig = px.imshow(pivot, text_auto='.1f', aspect="auto",
                    title="Тепловая карта: Средняя маржинальность по ABC-XYZ",
                    labels=dict(x="XYZ (маржинальность)", y="ABC (продажи)", color="Маржа %"),
                    color_continuous_scale="RdYlGn")
    fig.update_layout(height=400)
    return fig

# ============================================================================
# БЛОК 7: STREAMLIT ИНТЕРФЕЙС
# ============================================================================
def main():
    st.set_page_config(
        page_title=APP_NAME, 
        page_icon="📈", 
        layout="wide", 
        initial_sidebar_state="expanded"
    )
    
    if 'main_df' not in st.session_state:
        st.session_state.main_df = pd.DataFrame(columns=[
            'artikul', 'category', 'selling_price', 'cogs', 'daily_sales',
            'weight_kg', 'length_cm', 'width_cm', 'height_cm'
        ])
    
    # === БОКОВАЯ ПАНЕЛЬ ===
    st.sidebar.title("⚙️ Панель управления")
    st.sidebar.markdown(f"**{APP_NAME}**  \nВерсия: {APP_VERSION} © 2026", unsafe_allow_html=True)
    st.sidebar.markdown("---")
    
    # API доступы
    st.sidebar.subheader("🔐 API Доступы")
    st.sidebar.info("Ключи не хранятся в коде. Введите их здесь или используйте st.secrets.")
    
    api_key = st.sidebar.text_input(
        "API Key Яндекс Маркета", 
        type="password", 
        value=st.secrets.get("MARKET_API_KEY", "")
    )
    business_id = st.sidebar.text_input(
        "Business ID", 
        value=st.secrets.get("MARKET_BUSINESS_ID", "")
    )
    deepseek_key = st.sidebar.text_input(
        "API Key DeepSeek (fallback)", 
        type="password",
        value=st.secrets.get("DEEPSEEK_API_KEY", "")
    )
    
    use_api = st.sidebar.checkbox("🌐 Использовать API для тарифов", value=True, 
                                   help="Приоритет: API ЯМ → DeepSeek → локальная база")
    
    st.sidebar.markdown("---")
    
    # Настройки бизнеса
    st.sidebar.subheader("🏪 Настройки магазина")
    scheme_label = st.sidebar.selectbox(
        "Схема работы:",
        [s.value for s in YMScheme],
        help="FBS — со своего склада, FBY — со склада Маркета, Экспресс — быстрая доставка, DBS — своя доставка"
    )
    
    tax_label = st.sidebar.selectbox(
        "Налогообложение:",
        [t.label for t in TaxSystem],
        help="Выберите вашу систему налогообложения"
    )
    
    payment_freq = st.sidebar.selectbox(
        "Частота выплат:",
        ["Ежемесячно (1.0%)", "Раз в 2 недели (1.3%)", 
         "Еженедельно, 4 недели отсрочка (1.6%)",
         "Еженедельно, 2 недели отсрочка (2.3%)",
         "Еженедельно, 1 неделя отсрочка (2.8%)",
         "Ежедневно (3.3%)"],
        index=2,
        help="Влияет на ставку комиссии за перевод денег"
    )
    
    st.sidebar.markdown("---")
    
    # Навигация
    page = st.sidebar.radio("Выберите экран:", [
        "📊 Сводный Дашборд",
        "🔥 ABC-XYZ и 50+ Метрик",
        "💰 Рекомендованные цены",
        "📝 Калькулятор экономики",
        "🗂️ Управление категориями",
        "💾 Импорт / Экспорт",
        "📡 Синхронизация API"
    ])
    
    # === РАСЧЁТ ===
    df_hash = hashlib.md5(str(st.session_state.main_df.to_json()).encode()).hexdigest() if not st.session_state.main_df.empty else "empty"
    tariffs_snapshot = str(st.session_state.get('tariffs', {}))
    
    calculated_df = run_calculations_cached(df_hash, tax_label, tariffs_snapshot, scheme_label, payment_freq)
    
    # === ЭКРАНЫ ===
    if page == "📊 Сводный Дашборд":
        st.title("📊 Панель комплексной аналитики — Яндекс Маркет")
        
        if not calculated_df.empty:
            # KPI карточки
            c1, c2, c3, c4, c5 = st.columns(5)
            total_profit = calculated_df['gross_profit'].sum()
            
            c1.metric("Всего SKU", f"{len(calculated_df)}")
            c2.metric("Ср. маржинальность", f"{calculated_df['margin_percent'].mean():.2f}%")
            c3.metric("Общая выручка", format_number(calculated_df['selling_price'].sum(), " ₽"))
            c4.metric("ОБЩАЯ ПРИБЫЛЬ", format_number(total_profit, " ₽"), 
                      delta=f"{calculated_df['margin_percent'].mean():.1f}% маржа")
            c5.metric("Прибыльных / Убыточных", 
                      f"{(calculated_df['gross_profit'] > 0).sum()} / {(calculated_df['gross_profit'] < 0).sum()}")
            
            st.markdown("---")
            
            # Графики
            col1, col2 = st.columns(2)
            with col1:
                st.plotly_chart(render_waterfall_chart(calculated_df), use_container_width=True)
            with col2:
                st.plotly_chart(render_pareto_chart(calculated_df), use_container_width=True)
            
            col3, col4 = st.columns(2)
            with col3:
                avg_costs = calculated_df[['cogs', 'commission', 'delivery_to_customer', 
                                              'middle_mile_cost', 'sorting_cost', 'acquiring_cost',
                                              'tax_cost', 'return_cost', 'marketing_budget_per_unit',
                                              'warehouse_cost', 'pick_pack_cost', 'packaging_cost']].mean().reset_index()
                avg_costs.columns = ['Статья', 'Сумма']
                avg_costs = avg_costs[avg_costs['Сумма'] > 0]
                fig_tree = px.treemap(avg_costs, path=['Статья'], values='Сумма', 
                                      title="Средняя структура расходов на единицу",
                                      color='Сумма', color_continuous_scale='Reds')
                st.plotly_chart(fig_tree, use_container_width=True)
            
            with col4:
                st.plotly_chart(render_efficiency_heatmap(calculated_df), use_container_width=True)
            
            # Таблица ключевых метрик
            st.subheader("📋 Ключевые метрики SKU")
            display_cols = ['artikul', 'category', 'selling_price', 'cogs', 'gross_profit', 
                           'margin_percent', 'profitability_status', 'abc_xyz']
            display_cols = [c for c in display_cols if c in calculated_df.columns]
            st.dataframe(calculated_df[display_cols], use_container_width=True, hide_index=True)
            
        else:
            st.warning("Товарная матрица пуста. Загрузите файл на вкладке '💾 Импорт / Экспорт'.")
    
    elif page == "🔥 ABC-XYZ и 50+ Метрик":
        st.title("🔥 ABC-XYZ Матрица и 50+ Метрик")
        
        if not calculated_df.empty:
            st.plotly_chart(render_margin_scatter(calculated_df), use_container_width=True)
            
            st.subheader("50+ Метрик эффективности")
            st.dataframe(calculated_df, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            st.subheader("💾 Экспорт максимально информативных отчётов")
            
            if OPENPYXL_AVAILABLE:
                col_dl1, col_dl2 = st.columns(2)
                with col_dl1:
                    st.download_button(
                        label="⬇️ СКАЧАТЬ ПОЛНЫЙ ОТЧЁТ (.XLSX)",
                        data=UltimateExcelExporter.export_max_info(calculated_df, tax_label, scheme_label),
                        file_name=f"YM_UnitEconomics_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                        use_container_width=True,
                        type="primary"
                    )
                with col_dl2:
                    csv = calculated_df.to_csv(index=False).encode('utf-8-sig')
                    st.download_button(
                        label="⬇️ ЭКСПОРТ CSV",
                        data=csv,
                        file_name=f"YM_UnitEconomics_{datetime.now().strftime('%d_%m_%Y')}.csv",
                        use_container_width=True
                    )
            else:
                st.error("openpyxl не установлен. Установите: pip install openpyxl")
        else:
            st.info("Загрузите данные для формирования матрицы.")
    
    elif page == "💰 Рекомендованные цены":
        st.title("💰 Рекомендованные цены продажи")
        
        if not calculated_df.empty:
            price_cols = ['artikul', 'category', 'selling_price', 'cogs', 'gross_profit',
                          'margin_percent', 'rec_price_min', 'rec_price_10', 'rec_price_15',
                          'rec_price_20', 'rec_price_25', 'rec_price_30', 'profitability_status']
            price_cols = [c for c in price_cols if c in calculated_df.columns]
            
            st.dataframe(calculated_df[price_cols], use_container_width=True, hide_index=True)
            
            # График сравнения цен
            sample = calculated_df.head(20) if len(calculated_df) > 20 else calculated_df
            fig = go.Figure()
            fig.add_trace(go.Scatter(x=sample['artikul'], y=sample['selling_price'],
                                      mode='lines+markers', name='Текущая цена', line=dict(color='#1f77b4')))
            fig.add_trace(go.Scatter(x=sample['artikul'], y=sample['rec_price_min'],
                                      mode='lines+markers', name='Безубыточность', line=dict(color='#d62728', dash='dash')))
            fig.add_trace(go.Scatter(x=sample['artikul'], y=sample['rec_price_15'],
                                      mode='lines+markers', name='15% маржа', line=dict(color='#2ca02c')))
            fig.add_trace(go.Scatter(x=sample['artikul'], y=sample['rec_price_25'],
                                      mode='lines+markers', name='25% маржа', line=dict(color='#ff7f0e')))
            fig.update_layout(title="Сравнение текущих и рекомендованных цен (топ-20 SKU)",
                              xaxis_title="Артикул", yaxis_title="Цена, ₽",
                              height=500)
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Загрузите данные.")
    
    elif page == "📝 Калькулятор экономики":
        st.title("📝 Симулятор товарной матрицы")
        
        if not st.session_state.main_df.empty:
            edited_df = st.data_editor(
                st.session_state.main_df, 
                num_rows="dynamic", 
                use_container_width=True, 
                key="product_editor", 
                hide_index=True
            )
            if not edited_df.equals(st.session_state.main_df):
                st.session_state.main_df = edited_df
                st.rerun()
        else:
            st.info("Матрица пуста. Загрузите файл или введите данные.")
            dummy_df = pd.DataFrame([{
                'artikul': 'DEMO-001', 'category': 'автозапчасти', 
                'selling_price': 2500.0, 'cogs': 1200.0, 'daily_sales': 3,
                'weight_kg': 2.5, 'length_cm': 30, 'width_cm': 20, 'height_cm': 15,
                'volume_liters': 9.0, 'packaging_cost': 50.0, 'marketing_budget_per_unit': 100.0,
                'stock_depth_days': 30, 'first_mile_cost': 0, 'commission': 0,
                'return_cost': 0, 'warehouse_cost': 0, 'quantity_per_order': 1
            }])
            st.session_state.main_df = st.data_editor(dummy_df, num_rows="dynamic", 
                                                       use_container_width=True, hide_index=True)
    
    elif page == "🗂️ Управление категориями":
        st.title("🗂️ Индивидуальные тарифы категорий")
        
        tm = HybridTariffManager()
        
        # Текущие тарифы
        st.subheader("📋 Текущие тарифы в системе")
        st.dataframe(tm.to_dataframe(), use_container_width=True, hide_index=True)
        
        # Добавление новой категории
        with st.form("add_category_form"):
            st.subheader("➕ Добавить / изменить категорию")
            col1, col2, col3 = st.columns(3)
            with col1:
                new_cat = st.text_input("Название категории:")
            with col2:
                new_comm = st.number_input("Комиссия размещения (%):", value=12.0, min_value=0.0, max_value=100.0) / 100
            with col3:
                new_scheme = st.selectbox("Схема:", ["FBS", "FBY", "EXPRESS", "DBS"])
            
            col4, col5 = st.columns(2)
            with col4:
                new_delivery = st.number_input("Доставка покупателю (%):", value=4.5, min_value=0.0, max_value=100.0) / 100
            with col5:
                new_sorting = st.number_input("Обработка заказа (₽):", value=45.0, min_value=0.0)
            
            submitted = st.form_submit_button("💾 Сохранить тариф")
            if submitted and new_cat:
                tm.tariffs[new_cat.lower().strip()] = Tariff(
                    category=new_cat.lower().strip(),
                    commission_rate=new_comm,
                    sorting_cost=new_sorting,
                    delivery_rate=new_delivery,
                    source="Пользовательская база",
                    scheme=new_scheme
                )
                st.success(f"Категория '{new_cat}' добавлена!")
                st.rerun()
        
        # API тарифы
        st.markdown("---")
        st.subheader("🌐 Загрузить тарифы из API")
        if api_key and use_api:
            if st.button("🔄 Обновить тарифы через API Яндекс Маркета"):
                ym_api = YandexMarketAPI(api_key, business_id)
                campaigns = ym_api.get_campaigns()
                if campaigns:
                    st.success(f"Найдено {len(campaigns)} магазинов. Тарифы будут загружены при расчётах.")
                else:
                    st.warning("Не удалось получить список магазинов. Проверьте API ключ.")
        else:
            st.info("Введите API ключ Яндекс Маркета для загрузки актуальных тарифов.")
    
    elif page == "💾 Импорт / Экспорт":
        st.title("💾 Центр импорта и экспорта")
        
        st.markdown("### 📋 Загрузите ваш файл")
        st.info("Поддерживаются: CSV, XLSX, JSON. Автоматическое распознавание столбцов.")
        
        uploaded_file = st.file_uploader("Перетащите файл сюда", type=['csv', 'xlsx', 'json'])
        
        if uploaded_file is not None:
            try:
                bytes_data = uploaded_file.getvalue()
                raw_data = UniversalDataNormalizer.load_file_dynamically(
                    io.BytesIO(bytes_data), uploaded_file.name
                )
                processed_df = UniversalDataNormalizer.normalize_dataframe(raw_data)
                
                if not processed_df.empty:
                    st.session_state.main_df = processed_df
                    st.success(f"✅ Данные импортированы! Позиций: {len(processed_df)}")
                    st.dataframe(processed_df.head(10), use_container_width=True, hide_index=True)
                    
                    # Статистика импорта
                    st.markdown("#### 📊 Статистика импорта")
                    stat_cols = st.columns(4)
                    stat_cols[0].metric("SKU", len(processed_df))
                    stat_cols[1].metric("Категорий", processed_df['category'].nunique())
                    stat_cols[2].metric("Ср. цена", f"{processed_df['selling_price'].mean():.0f} ₽")
                    stat_cols[3].metric("Ср. себестоимость", f"{processed_df['cogs'].mean():.0f} ₽")
                else:
                    st.error("В файле отсутствуют данные.")
            except Exception as e:
                st.error(f"Ошибка импорта: {str(e)}")
    
    elif page == "📡 Синхронизация API":
        st.title("📡 API-шлюз Яндекс Маркета")
        
        if not api_key:
            st.warning("⚠️ Введите API Key в боковой панели.")
        else:
            ym_api = YandexMarketAPI(api_key, business_id)
            
            # Статус API
            st.subheader("🔍 Статус подключения")
            campaigns = ym_api.get_campaigns()
            if campaigns:
                st.success(f"✅ Подключено! Магазинов: {len(campaigns)}")
                camp_df = pd.DataFrame([{
                    'ID': c.get('id'), 
                    'Название': c.get('domain', 'N/A'),
                    'Статус': c.get('state', 'N/A')
                } for c in campaigns])
                st.dataframe(camp_df, use_container_width=True, hide_index=True)
            else:
                st.error("❌ Не удалось подключиться к API. Проверьте ключ.")
            
            # Расчёт тарифов через API
            st.markdown("---")
            st.subheader("🧮 Расчёт тарифов через API")
            
            if not calculated_df.empty and campaigns:
                campaign_id = st.selectbox("Выберите магазин:", 
                                           [c['id'] for c in campaigns],
                                           format_func=lambda x: f"ID: {x}")
                
                if st.button("🚀 РАССЧИТАТЬ ТАРИФЫ ЧЕРЕЗ API", type="primary"):
                    with st.spinner("Запрос к API Яндекс Маркета..."):
                        sample = calculated_df.head(10)
                        offers = []
                        for _, row in sample.iterrows():
                            offers.append({
                                "categoryId": 0,
                                "price": float(row['selling_price']),
                                "length": float(row.get('length_cm', 10)),
                                "width": float(row.get('width_cm', 10)),
                                "height": float(row.get('height_cm', 10)),
                                "weight": float(row.get('weight_kg', 1)),
                                "quantity": int(row.get('quantity_per_order', 1))
                            })
                        
                        scheme = scheme_label.split(" ")[0]
                        result = ym_api.calculate_tariffs(offers, campaign_id=campaign_id, 
                                                           selling_program=scheme)
                        
                        if result:
                            st.success(f"✅ Получены тарифы для {len(result)} товаров")
                            st.json(result[:2])
                        else:
                            st.error("API вернул пустой результат.")
            
            # Выгрузка цен
            st.markdown("---")
            st.subheader("📤 Выгрузка цен на Маркет")
            
            if not calculated_df.empty:
                selected_strategy = st.selectbox(
                    "Стратегия цены:",
                    ["Текущая цена", "Минимальная (безубыточность)", 
                     "Оптимальная (15% маржа)", "Агрессивная (25% маржа)"]
                )
                
                if st.button("🚀 ВЫГРУЗИТЬ ЦЕНЫ НА МАРКЕТПЛЕЙС", type="primary"):
                    price_data = []
                    for _, row in calculated_df.iterrows():
                        if selected_strategy == "Текущая цена":
                            target = row.get('selling_price', 0)
                        elif selected_strategy == "Минимальная (безубыточность)":
                            target = row.get('rec_price_min', 0)
                        elif selected_strategy == "Оптимальная (15% маржа)":
                            target = row.get('rec_price_15', 0)
                        else:
                            target = row.get('rec_price_25', 0)
                        price_data.append({
                            'artikul': row.get('artikul', ''),
                            'new_price': max(target, row.get('cogs', 0) * 1.05)
                        })
                    
                    with st.spinner("Подготовка данных..."):
                        st.success(f"✅ Готово к выгрузке {len(price_data)} позиций")
                        st.dataframe(pd.DataFrame(price_data).head(10), use_container_width=True, hide_index=True)
                        st.info("Реальная выгрузка требует вызова POST v2/businesses/{businessId}/offer-prices/updates")

if __name__ == "__main__":
    main()
