#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
🚀 ULTIMATE UNIT ECONOMICS FOR YANDEX MARKET v21.1 - UNIVERSAL EXCEL COMPAT
============================================================================
Точные подсчёты. Тарифы только из API или загружаемого справочника.
Excel-выгрузка содержит живые формулы (IFERROR + VLOOKUP), совместимые 
с Excel 2010+ и Google Таблицами, а также встроенные дашборды.
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
from typing import Dict, List, Optional

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import DataBarRule, ColorScaleRule, FormulaRule
    from openpyxl.chart import BarChart, PieChart, Reference
    from openpyxl.chart.label import DataLabelList
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('YandexMarketUnitEconomics')

APP_VERSION = "21.1.0"
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

def fix_double_utf8(text: str) -> str:
    if not isinstance(text, str) or not text: return text
    for source_enc, target_enc in [('cp1251', 'utf-8'), ('latin1', 'utf-8')]:
        try:
            fixed = text.encode(source_enc).decode(target_enc)
            if fixed and 'Р' not in fixed[:2]: return fixed
        except Exception: continue
    return text

def format_number(num: float, suffix='') -> str:
    if pd.isna(num): return "0"
    abs_num = abs(num)
    sign = "-" if num < 0 else ""
    for unit in ['', 'K', 'M', 'B']:
        if abs_num < 1000.0: return f"{sign}{abs_num:3.1f}{unit}{suffix}".strip()
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
    
    def __init__(self, label, rate, base, min_rate):
        self.label = label; self.rate = rate; self.base = base; self.min_rate = min_rate
    
    @classmethod
    def by_label(cls, label):
        for item in cls:
            if item.label == label: return item
        return cls.USN_6

class YMScheme(Enum):
    FBS = "FBS (склад продавца)"
    FBY = "FBY (склад Маркета)"
    EXPRESS = "Экспресс"
    DBS = "DBS (доставка продавца)"

class Tariff:
    def __init__(self, category: str, commission_rate: float = 0.15, min_commission: float = 0.0,
                 sorting_cost: float = 45.0, delivery_rate: float = 0.045, delivery_min: float = 60.0, 
                 delivery_max: float = 500.0, acquiring_transfer_rate: float = 0.016, acquiring_sku_cost: float = 0.12,
                 return_rate: float = 0.05, return_processing: float = 15.0, storage_fee_per_day: float = 0.50,
                 special_tariff_rate: float = 0.42, source: str = "Неизвестно", scheme: str = "FBS"):
        self.category = category
        self.commission_rate = commission_rate
        self.min_commission = min_commission
        self.sorting_cost = sorting_cost
        self.delivery_rate = delivery_rate
        self.delivery_min = delivery_min
        self.delivery_max = delivery_max
        self.acquiring_transfer_rate = acquiring_transfer_rate
        self.acquiring_sku_cost = acquiring_sku_cost
        self.return_rate = return_rate
        self.return_processing = return_processing
        self.storage_fee_per_day = storage_fee_per_day
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
        self.headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json", "Accept": "application/json"}
        if business_id: self.headers["X-Business-Id"] = business_id
    
    def get_campaigns(self) -> List[Dict]:
        try:
            resp = requests.get(f"{self.BASE_URL}/v2/campaigns", headers=self.headers, timeout=15)
            resp.raise_for_status()
            return resp.json().get("campaigns", [])
        except Exception as e:
            logger.error(f"Ошибка API ЯМ: {e}")
            return []
    
    def calculate_tariffs(self, offers: List[Dict], campaign_id: int = None, selling_program: str = "FBS") -> List[Dict]:
        try:
            payload = {"parameters": {"sellingProgram": selling_program, "frequency": "WEEKLY", "paymentDelayWeeks": 4, "currency": "RUR"}, "offers": offers}
            if campaign_id:
                payload["parameters"]["campaignId"] = campaign_id
                del payload["parameters"]["sellingProgram"]
            resp = requests.post(f"{self.BASE_URL}/v2/tariffs/calculate", headers=self.headers, json=payload, timeout=30)
            resp.raise_for_status()
            return resp.json().get("result", {}).get("offers", [])
        except Exception as e:
            logger.error(f"Ошибка расчета тарифов API: {e}")
            return []

class DeepSeekAPI:
    BASE_URL = "https://api.deepseek.com"
    def __init__(self, api_key: str):
        self.api_key = api_key
        self.headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
    
    def analyze_tariffs(self, category_name: str, scheme: str = "FBS") -> Dict:
        try:
            prompt = f"""Ты — эксперт по тарифам Яндекс Маркета. Для категории "{category_name}" и схемы {scheme} 
укажи актуальные тарифы на 2026 год СТРОГО в формате JSON без markdown-оберток:
{{"commission_rate": 0.15, "sorting_cost": 45, "delivery_rate": 0.045, "delivery_min": 60, "delivery_max": 500, "acquiring_transfer_rate": 0.016, "acquiring_sku_cost": 0.12, "return_rate": 0.05, "return_processing": 15, "storage_fee_per_day": 0.5, "special_tariff_rate": 0.42}}"""
            resp = requests.post(f"{self.BASE_URL}/chat/completions", headers=self.headers,
                json={"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}], "temperature": 0.1}, timeout=15)
            resp.raise_for_status()
            content = resp.json()["choices"][0]["message"]["content"]
            json_match = re.search(r'\{.*\}', content, re.DOTALL)
            return json.loads(json_match.group()) if json_match else {}
        except Exception as e:
            logger.error(f"Ошибка DeepSeek: {e}")
            return {}

class HybridTariffManager:
    def __init__(self):
        if 'tariffs' not in st.session_state: st.session_state.tariffs = {}
        if 'ym_api_cache' not in st.session_state: st.session_state.ym_api_cache = {}

    @property
    def tariffs(self): return st.session_state.tariffs

    def load_tariffs_from_file(self, df: pd.DataFrame):
        req_cols = ['category', 'commission_rate']
        if not all(col in df.columns for col in req_cols):
            raise ValueError("Файл тарифов должен содержать минимум: category, commission_rate")
        for _, row in df.iterrows():
            cat = str(row['category']).lower().strip()
            self.tariffs[cat] = Tariff(
                category=cat, commission_rate=float(row.get('commission_rate', 0.15)),
                min_commission=float(row.get('min_commission', 0)), sorting_cost=float(row.get('sorting_cost', 45)),
                delivery_rate=float(row.get('delivery_rate', 0.045)), delivery_min=float(row.get('delivery_min', 60)),
                delivery_max=float(row.get('delivery_max', 500)), acquiring_transfer_rate=float(row.get('acquiring_transfer_rate', 0.016)),
                acquiring_sku_cost=float(row.get('acquiring_sku_cost', 0.12)), return_rate=float(row.get('return_rate', 0.05)),
                return_processing=float(row.get('return_processing', 15)), storage_fee_per_day=float(row.get('storage_fee_per_day', 0.5)),
                special_tariff_rate=float(row.get('special_tariff_rate', 0.42)), source="Загружено пользователем", scheme=row.get('scheme', 'FBS')
            )

    def get_best_tariff(self, category_name: str, scheme: str, ym_api: YandexMarketAPI = None, deepseek_api: DeepSeekAPI = None, use_api: bool = True) -> Tariff:
        cat_clean = str(category_name).lower().strip()
        cache_key = f"{cat_clean}_{scheme}"
        if cache_key in st.session_state.ym_api_cache: return st.session_state.ym_api_cache[cache_key]
        
        if use_api and ym_api and ym_api.api_key:
            try:
                result = ym_api.calculate_tariffs([{"categoryId": 0, "price": 1000, "length": 10, "width": 10, "height": 10, "weight": 1, "quantity": 1}], selling_program=scheme)
                if result and len(result) > 0:
                    t = self._parse_ym_tariffs(result[0].get("tariffs", []), cat_clean, scheme)
                    if t: st.session_state.ym_api_cache[cache_key] = t; return t
            except Exception as e: logger.warning(f"API ЯМ сбой для {cat_clean}: {e}")
        
        if use_api and deepseek_api and deepseek_api.api_key:
            try:
                ds_result = deepseek_api.analyze_tariffs(cat_clean, scheme)
                if ds_result and 'commission_rate' in ds_result:
                    t = Tariff(category=cat_clean, commission_rate=ds_result.get('commission_rate', 0.15), sorting_cost=ds_result.get('sorting_cost', 45),
                               delivery_rate=ds_result.get('delivery_rate', 0.045), delivery_min=ds_result.get('delivery_min', 60), delivery_max=ds_result.get('delivery_max', 500),
                               acquiring_transfer_rate=ds_result.get('acquiring_transfer_rate', 0.016), acquiring_sku_cost=ds_result.get('acquiring_sku_cost', 0.12),
                               return_rate=ds_result.get('return_rate', 0.05), return_processing=ds_result.get('return_processing', 15),
                               storage_fee_per_day=ds_result.get('storage_fee_per_day', 0.5), special_tariff_rate=ds_result.get('special_tariff_rate', 0.42),
                               source=f"DeepSeek AI ({scheme})", scheme=scheme)
                    st.session_state.ym_api_cache[cache_key] = t; return t
            except Exception as e: logger.warning(f"DeepSeek сбой для {cat_clean}: {e}")
        
        if cat_clean in self.tariffs:
            t = self.tariffs[cat_clean]; t.scheme = scheme; return t
            
        logger.warning(f"Тариф для {cat_clean} не найден. Применен базовый фоллбэк 15%.")
        t = Tariff(category=cat_clean, commission_rate=0.15, source="⚠️ БАЗОВЫЙ ФОЛЛБЭК (ТРЕБУЕТ ПРОВЕРКИ)", scheme=scheme)
        st.session_state.ym_api_cache[cache_key] = t
        return t

    def _parse_ym_tariffs(self, tariffs_data: List[Dict], category: str, scheme: str) -> Optional[Tariff]:
        if not tariffs_data: return None
        comm_rate, sort_cost, del_rate, acq_rate = 0.15, 45.0, 0.045, 0.016
        for t in tariffs_data:
            t_type, amount = t.get("type", ""), t.get("amount", 0)
            params = {p.get("name", "").lower(): p.get("value", "") for p in t.get("parameters", [])}
            if t_type == "FEE" and params.get("valuetype") == "relative": comm_rate = amount / 100.0
            elif t_type == "SORTING": sort_cost = amount
            elif t_type == "DELIVERY_TO_CUSTOMER" and params.get("valuetype") == "relative": del_rate = amount / 100.0
            elif t_type == "PAYMENT_TRANSFER" and params.get("valuetype") == "relative": acq_rate = amount / 100.0
        return Tariff(category=category, commission_rate=comm_rate, sorting_cost=sort_cost, delivery_rate=del_rate, 
                      acquiring_transfer_rate=acq_rate, source=f"API Яндекс Маркета ({scheme})", scheme=scheme)
    
    def to_dataframe(self) -> pd.DataFrame:
        if not self.tariffs: return pd.DataFrame(columns=['Категория', 'Комиссия, %', 'Источник данных'])
        return pd.DataFrame([{'Категория': k, 'Комиссия, %': round(t.commission_rate * 100, 2), 'Мин. комиссия, ₽': t.min_commission,
            'Сортировка, ₽': t.sorting_cost, 'Доставка %': round(t.delivery_rate * 100, 2), 'Доставка мин, ₽': t.delivery_min,
            'Доставка макс, ₽': t.delivery_max, 'Эквайринг перевод, %': round(t.acquiring_transfer_rate * 100, 2),
            'Эквайринг SKU, ₽': t.acquiring_sku_cost, 'Возвраты, %': round(t.return_rate * 100, 2),
            'Обработка возврата, ₽': t.return_processing, 'Хранение день, ₽': t.storage_fee_per_day,
            'Спецтариф <=300₽, %': round(t.special_tariff_rate * 100, 2), 'Схема': t.scheme, 'Источник данных': t.source} for k, t in self.tariffs.items()])

# ============================================================================
# БЛОК 3: ВЕКТОРИЗОВАННЫЙ ФИНАНСОВЫЙ ДВИЖОК
# ============================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def run_calculations_cached(df_hash: str, tax_label: str, scheme_label: str, payment_frequency: str, tariffs_snapshot_hash: str) -> pd.DataFrame:
    df = st.session_state.main_df.copy()
    if df.empty: return df
    
    tax_system = TaxSystem.by_label(tax_label)
    scheme = scheme_label.split(" ")[0]
    payment_rates = {"Ежемесячно (1.0%)": 0.01, "Раз в 2 недели (1.3%)": 0.013, "Еженедельно, 4 нед. (1.6%)": 0.016, "Ежедневно (3.3%)": 0.033}
    p_transfer_rate = payment_rates.get(payment_frequency, 0.016)
    
    manager = HybridTariffManager()
    for col in ['artikul', 'category']:
        if col in df.columns: df[col] = df[col].astype(str).apply(fix_double_utf8)
    
    req_cols = {'selling_price': 0.0, 'cogs': 0.0, 'weight_kg': 0.0, 'length_cm': 0.0, 'width_cm': 0.0, 'height_cm': 0.0,
                'packaging_cost': 0.0, 'marketing_budget_per_unit': 0.0, 'daily_sales': 0.0, 'stock_depth_days': 0.0,
                'first_mile_cost': 0.0, 'commission': 0.0, 'return_cost': 0.0, 'warehouse_cost': 0.0, 'volume_liters': 0.0, 'quantity_per_order': 1.0}
    for col, default in req_cols.items():
        if col not in df.columns: df[col] = default
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(default)
    
    comm_rates, sort_costs, del_rates, del_mins, del_maxs, acq_rates, acq_skus, ret_rates, ret_proc, stor_fees, spec_rates = [], [], [], [], [], [], [], [], [], [], []
    for cat in df.get('category', ['default'] * len(df)):
        t = manager.get_best_tariff(cat, scheme)
        comm_rates.append(t.commission_rate); sort_costs.append(t.sorting_cost); del_rates.append(t.delivery_rate)
        del_mins.append(t.delivery_min); del_maxs.append(t.delivery_max); acq_rates.append(p_transfer_rate)
        acq_skus.append(t.acquiring_sku_cost); ret_rates.append(t.return_rate); ret_proc.append(t.return_processing)
        stor_fees.append(t.storage_fee_per_day); spec_rates.append(t.special_tariff_rate)
    
    comm_rates, del_rates, acq_rates, ret_rates, spec_rates = map(np.array, [comm_rates, del_rates, acq_rates, ret_rates, spec_rates])
    
    vol_weight = (df['length_cm'] * df['width_cm'] * df['height_cm']) / 5000.0
    df['billable_weight'] = np.ceil(np.maximum(df['weight_kg'], vol_weight) * 2) / 2
    df['is_special_tariff'] = (df['selling_price'] <= 300) & (df['volume_liters'] <= 5)
    
    df['commission'] = np.where(df['is_special_tariff'], df['selling_price'] * spec_rates, np.maximum(df['selling_price'] * comm_rates, 0))
    df['delivery_to_customer'] = np.where(df['is_special_tariff'], 0.0, np.clip(df['selling_price'] * del_rates, np.array(del_mins), np.array(del_maxs)))
    df['middle_mile_cost'] = np.where(df['is_special_tariff'], 0.0, np.where(df['billable_weight'] <= 4, 100, np.where(df['billable_weight'] <= 10, 300, 600)))
    df['sorting_cost'] = np.where(df['is_special_tariff'], 0.0, np.where(scheme == 'FBS', np.array(sort_costs), 0.0))
    df['acquiring_sku_cost'] = np.where(df['quantity_per_order'] > 0, np.array(acq_skus) / df['quantity_per_order'], np.array(acq_skus))
    df['acquiring_transfer_cost'] = df['selling_price'] * acq_rates
    df['acquiring_cost'] = df['acquiring_sku_cost'] + df['acquiring_transfer_cost']
    df['return_processing_cost'] = np.where(df['is_special_tariff'], 0.0, np.array(ret_proc))
    df['return_delivery_cost'] = np.where(df['is_special_tariff'], 0.0, df['middle_mile_cost'] * ret_rates)
    df['return_cost'] = df['return_processing_cost'] + df['return_delivery_cost']
    df['pick_pack_cost'] = 35.0
    df['warehouse_cost'] = np.where(df['warehouse_cost'] == 0, (df['stock_depth_days'] * df['daily_sales']) * np.array(stor_fees), df['warehouse_cost'])
    
    df['fixed_operational_costs'] = df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + df['packaging_cost'] + df['return_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    df['marketplace_fees'] = df['commission'] + df['delivery_to_customer'] + df['middle_mile_cost'] + df['sorting_cost'] + df['acquiring_cost']
    df['pre_tax_expenses'] = df['fixed_operational_costs'] + df['marketplace_fees']
    
    if tax_system.base == "revenue": df['tax_cost'] = df['selling_price'] * tax_system.rate
    else:
        pre_tax_profit = (df['selling_price'] - (df['selling_price'] * 0.20 / 1.20)) - df['pre_tax_expenses'] if tax_system.base == "profit_vat" else df['selling_price'] - df['pre_tax_expenses']
        df['tax_cost'] = np.maximum(pre_tax_profit, 0) * tax_system.rate
    
    df['total_expenses'] = df['pre_tax_expenses'] + df['tax_cost']
    df['gross_profit'] = df['selling_price'] - df['total_expenses']
    df['margin_percent'] = np.where(df['selling_price'] > 0, (df['gross_profit'] / df['selling_price']) * 100, 0.0)
    
    var_fees = np.where(df['is_special_tariff'], spec_rates + acq_rates + (tax_system.rate if tax_system.base=="revenue" else 0), comm_rates + del_rates + acq_rates + (tax_system.rate if tax_system.base=="revenue" else 0))
    denom = np.where((1.0 - var_fees) <= 0.01, 0.5, 1.0 - var_fees)
    fixed_no_return = df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + df['packaging_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    df['rec_price_min'] = fixed_no_return / denom
    df['rec_price_15'] = fixed_no_return / (denom - 0.15)
    df['rec_price_25'] = fixed_no_return / (denom - 0.25)
    
    df['variable_costs'] = df['commission'] + df['delivery_to_customer'] + df['middle_mile_cost'] + df['sorting_cost'] + df['acquiring_cost'] + df['return_cost']
    df['fixed_costs'] = df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + df['packaging_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    df['contribution_margin'] = df['selling_price'] - df['variable_costs']
    df['roi_percent'] = np.where(df['cogs'] > 0, ((df['gross_profit'] / df['cogs']) * 100), 0.0)
    df['break_even_units'] = np.where(df['contribution_margin'] > 0, df['fixed_costs'] / df['contribution_margin'], 0.0)
    df['abc_category'] = np.where(df['daily_sales'] >= 10, 'A', np.where(df['daily_sales'] >= 3, 'B', 'C'))
    df['xyz_category'] = np.where(df['margin_percent'] >= 20, 'X', np.where(df['margin_percent'] >= 10, 'Y', 'Z'))
    df['abc_xyz'] = df['abc_category'] + df['xyz_category']
    df['profitability_status'] = np.where(df['gross_profit'] > 0, np.where(df['margin_percent'] >= 20, 'Высокомаржинальный', 'Низкомаржинальный'), 'Убыточный')
    
    for col in ['commission', 'delivery_to_customer', 'middle_mile_cost', 'sorting_cost', 'acquiring_cost', 'return_cost', 'gross_profit', 'total_expenses', 'rec_price_min', 'rec_price_15', 'rec_price_25']:
        if col in df.columns: df[col] = df[col].apply(money_round)
    for col in ['margin_percent', 'roi_percent']:
        if col in df.columns: df[col] = df[col].apply(percent_round)
    
    return df

# ============================================================================
# БЛОК 4: ЭКСПОРТЁР С УНИВЕРСАЛЬНЫМИ ФОРМУЛАМИ (VLOOKUP)
# ============================================================================
class UltimateExcelExporter:
    @staticmethod
    def _get_col_letter(idx: int) -> str:
        res = ""
        while idx >= 0: res = chr(idx % 26 + 65) + res; idx = idx // 26 - 1
        return res

    @staticmethod
    def export_max_info(df: pd.DataFrame, tax_label: str, scheme_label: str, tariff_manager: HybridTariffManager) -> bytes:
        if not OPENPYXL_AVAILABLE or df.empty: return b""
        wb = Workbook()
        
        # === 1. ЛИСТ СПРАВОЧНИКА ТАРИФОВ (15 столбцов, жесткая структура для VLOOKUP) ===
        ws_tariffs = wb.create_sheet("Тарифы_Справочник", 0)
        tariff_df = tariff_manager.to_dataframe()
        t_cols = list(tariff_df.columns)
        for c_idx, c_name in enumerate(t_cols, 1):
            cell = ws_tariffs.cell(1, c_idx, c_name)
            cell.font = Font(bold=True, color="FFFFFF", size=11); cell.fill = PatternFill(start_color="1F4E78", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
        for r_idx, row_data in enumerate(tariff_df.itertuples(index=False), 2):
            for c_idx, value in enumerate(row_data, 1): ws_tariffs.cell(r_idx, c_idx, value)
        for c_idx in range(1, len(t_cols) + 1): ws_tariffs.column_dimensions[UltimateExcelExporter._get_col_letter(c_idx - 1)].width = 18

        # === 2. ДЕТАЛЬНЫЙ РАСЧЁТ С УНИВЕРСАЛЬНЫМИ ФОРМУЛАМИ ===
        ws_detail = wb.create_sheet("Детальный_Расчет")
        priority_cols = ['artikul', 'category', 'selling_price', 'cogs', 'weight_kg', 'billable_weight', 'commission', 
                         'delivery_to_customer', 'middle_mile_cost', 'sorting_cost', 'acquiring_cost', 'return_cost', 
                         'fixed_operational_costs', 'marketplace_fees', 'total_expenses', 'gross_profit', 'margin_percent', 
                         'rec_price_min', 'rec_price_15', 'rec_price_25', 'profitability_status', 'abc_xyz']
        cols = [c for c in priority_cols if c in df.columns] + [c for c in df.columns if c not in priority_cols]
        
        for c_idx, c_name in enumerate(cols, 1):
            cell = ws_detail.cell(1, c_idx, c_name)
            cell.font = Font(bold=True, color="FFFFFF", size=10); cell.fill = PatternFill(start_color="2E75B6", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)
            cell.border = Border(left=Side(style="thin"), right=Side(style="thin"), top=Side(style="thin"), bottom=Side(style="thin"))
        ws_detail.auto_filter.ref = f"A1:{UltimateExcelExporter._get_col_letter(len(cols)-1)}{len(df)+1}"
        ws_detail.freeze_panes = 'B2'
        
        idx_map = {col: i for i, col in enumerate(cols)}
        def cell_ref(col_name: str, row_num: int) -> str:
            return f"{UltimateExcelExporter._get_col_letter(idx_map[col_name])}{row_num}" if col_name in idx_map else ""

        # Хелпер для генерации универсальной формулы VLOOKUP + IFERROR
        def vlookup_formula(lookup_val_cell: str, col_index: int, default_val: float) -> str:
            # $A:$O охватывает все 15 столбцов справочника. FALSE обеспечивает точное совпадение.
            return f"=IFERROR(VLOOKUP({lookup_val_cell}, 'Тарифы_Справочник'!$A:$O, {col_index}, FALSE), {default_val})"

        for r_idx, row_data in enumerate(df.itertuples(index=False), 2):
            cat_cell = cell_ref('category', r_idx)
            price_cell = cell_ref('selling_price', r_idx)
            
            for c_idx, value in enumerate(row_data, 1):
                cell = ws_detail.cell(r_idx, c_idx, value)
                cell.border = Border(bottom=Side(style="thin", color="E0E0E0"))
                col_name = cols[c_idx - 1]
                if isinstance(value, (int, float)):
                    cell.number_format = '0.00"%"' if 'percent' in col_name else ('0' if col_name in ['daily_sales', 'break_even_units'] else '#,##0.00')

            # ЖИВЫЕ ФОРМУЛЫ (Совместимы с Excel 2010+ и Google Sheets)
            # Индексы столбцов в 'Тарифы_Справочник': 2=Комиссия%, 4=Сортировка, 5=Доставка%, 8=Эквайринг%
            if 'commission' in idx_map and cat_cell and price_cell:
                ws_detail.cell(r_idx, idx_map['commission'] + 1).value = f"={price_cell} * ({vlookup_formula(cat_cell, 2, 15)} / 100)"
            if 'sorting_cost' in idx_map and cat_cell:
                ws_detail.cell(r_idx, idx_map['sorting_cost'] + 1).value = vlookup_formula(cat_cell, 4, 45)
            if 'delivery_to_customer' in idx_map and cat_cell and price_cell:
                ws_detail.cell(r_idx, idx_map['delivery_to_customer'] + 1).value = f"={price_cell} * ({vlookup_formula(cat_cell, 5, 4.5)} / 100)"
            if 'acquiring_cost' in idx_map and cat_cell and price_cell:
                ws_detail.cell(r_idx, idx_map['acquiring_cost'] + 1).value = f"={price_cell} * ({vlookup_formula(cat_cell, 8, 1.6)} / 100)"
            
            # Агрегирующие формулы
            if all(c in idx_map for c in ['commission', 'delivery_to_customer', 'middle_mile_cost', 'sorting_cost', 'acquiring_cost']) and 'marketplace_fees' in idx_map:
                ws_detail.cell(r_idx, idx_map['marketplace_fees'] + 1).value = f"={cell_ref('commission', r_idx)} + {cell_ref('delivery_to_customer', r_idx)} + {cell_ref('middle_mile_cost', r_idx)} + {cell_ref('sorting_cost', r_idx)} + {cell_ref('acquiring_cost', r_idx)}"
            if all(c in idx_map for c in ['selling_price', 'total_expenses', 'gross_profit']):
                ws_detail.cell(r_idx, idx_map['gross_profit'] + 1).value = f"={cell_ref('selling_price', r_idx)} - {cell_ref('total_expenses', r_idx)}"
            if all(c in idx_map for c in ['gross_profit', 'selling_price', 'margin_percent']):
                cell = ws_detail.cell(r_idx, idx_map['margin_percent'] + 1)
                cell.value = f"=IF({cell_ref('selling_price', r_idx)}>0, {cell_ref('gross_profit', r_idx)} / {cell_ref('selling_price', r_idx)}, 0)"
                cell.number_format = '0.00%'

        # Условное форматирование
        if 'gross_profit' in idx_map:
            col_l = UltimateExcelExporter._get_col_letter(idx_map['gross_profit'])
            ws_detail.conditional_formatting.add(f'{col_l}2:{col_l}{len(df)+1}', FormulaRule(formula=[f'{col_l}2>=0'], fill=PatternFill(start_color="C6EFCE"), font=Font(color="006100", bold=True)))
            ws_detail.conditional_formatting.add(f'{col_l}2:{col_l}{len(df)+1}', FormulaRule(formula=[f'{col_l}2<0'], fill=PatternFill(start_color="FFC7CE"), font=Font(color="9C0006", bold=True)))
        if 'margin_percent' in idx_map:
            col_l = UltimateExcelExporter._get_col_letter(idx_map['margin_percent'])
            ws_detail.conditional_formatting.add(f'{col_l}2:{col_l}{len(df)+1}', DataBarRule(start_type='min', end_type='max', color="638EC6", showValue=True))
        
        for c_idx, c_name in enumerate(cols, 1):
            ws_detail.column_dimensions[UltimateExcelExporter._get_col_letter(c_idx - 1)].width = 22 if c_name in ['artikul', 'category', 'profitability_status'] else (14 if 'percent' in c_name else 18)

        # === 3. ДАШБОРД-СВОДКА С ВИЗУАЛИЗАЦИЕЙ ===
        ws_dash = wb.create_sheet("Дашборд_Сводка", 1)
        ws_dash.merge_cells('A1:D1')
        cell = ws_dash.cell(1, 1, "СВОДНЫЙ ФИНАНСОВЫЙ ДАШБОРД"); cell.font = Font(size=16, bold=True, color="1F4E78"); cell.alignment = Alignment(horizontal="center")
        
        metrics = [("Всего SKU", len(df)), ("Общая выручка", df['selling_price'].sum()), ("ОБЩАЯ ПРИБЫЛЬ", df['gross_profit'].sum()), ("Средняя маржа %", df['margin_percent'].mean())]
        for r_idx, (label, val) in enumerate(metrics, 3):
            ws_dash.cell(r_idx, 1, label).font = Font(bold=True)
            c = ws_dash.cell(r_idx, 2, val); c.number_format = '0.00"%"' if '%' in label else '#,##0.00 "₽"'; c.font = Font(bold=True, color="1F4E78")

        expense_labels = ["Себестоимость", "Комиссия", "Доставка", "Ср. миля", "Эквайринг", "Налоги"]
        expense_vals = [df['cogs'].sum(), df['commission'].sum(), df['delivery_to_customer'].sum(), df['middle_mile_cost'].sum(), df['acquiring_cost'].sum(), df['tax_cost'].sum()]
        ws_dash.cell(10, 1, "Структура расходов"); ws_dash.cell(10, 1).font = Font(bold=True, size=12)
        for i, (lbl, val) in enumerate(zip(expense_labels, expense_vals), 11):
            ws_dash.cell(i, 1, lbl); ws_dash.cell(i, 2, val)
        
        pie = PieChart()
        pie.title = "Структура расходов"
        labels = Reference(ws_dash, min_col=1, min_row=11, max_row=16)
        data = Reference(ws_dash, min_col=2, min_row=11, max_row=16)
        pie.add_data(data, titles_from_data=False); pie.set_categories(labels)
        pie.dataLabels = DataLabelList(); pie.dataLabels.showPercent = True
        ws_dash.add_chart(pie, "D10")

        # === 4. ABC-XYZ АНАЛИЗ ===
        ws_abc = wb.create_sheet("ABC_XYZ")
        abc_data = df.groupby('abc_xyz').agg({'artikul': 'count', 'selling_price': 'sum', 'gross_profit': 'sum', 'margin_percent': 'mean'}).reset_index()
        abc_data.columns = ['ABC-XYZ', 'Кол-во SKU', 'Выручка', 'Прибыль', 'Ср. маржа %']
        for c_idx, c_name in enumerate(abc_data.columns, 1):
            cell = ws_abc.cell(1, c_idx, c_name); cell.font = Font(bold=True, color="FFFFFF"); cell.fill = PatternFill(start_color="2E75B6", fill_type="solid")
        for r_idx, row_data in enumerate(abc_data.itertuples(index=False), 2):
            for c_idx, val in enumerate(row_data, 1):
                cell = ws_abc.cell(r_idx, c_idx, val)
                cell.number_format = '#,##0.00' if c_idx > 2 else '0.00"%"' if c_idx == 5 else '0'
        
        bar = BarChart()
        bar.title = "Прибыль по ABC-XYZ сегментам"
        bar.x_axis.title = "Сегмент"
        bar.y_axis.title = "Прибыль, ₽"
        cats = Reference(ws_abc, min_col=1, min_row=2, max_row=len(abc_data)+1)
        vals = Reference(ws_abc, min_col=4, min_row=1, max_row=len(abc_data)+1)
        bar.add_data(vals, titles_from_data=True); bar.set_categories(cats)
        ws_abc.add_chart(bar, "G2")

        out = io.BytesIO()
        wb.save(out); out.seek(0)
        return out.getvalue()

# ============================================================================
# БЛОК 5: НОРМАЛИЗАТОР ДАННЫХ
# ============================================================================
class UniversalDataNormalizer:
    COLUMN_MAPPING = {
        'artikul': ['artikul', 'артикул', 'sku', 'offer_id', 'id'],
        'category': ['category', 'категория', 'группа', 'предмет'],
        'selling_price': ['selling_price', 'цена продажи', 'цена', 'price'],
        'cogs': ['cogs', 'себестоимость', 'закупка', 'cost'],
        'daily_sales': ['daily_sales', 'заказы, шт.', 'продажи, шт.', 'quantity'],
        'weight_kg': ['weight_kg', 'вес', 'weight', 'вес кг'],
        'length_cm': ['length_cm', 'длина', 'length'], 'width_cm': ['width_cm', 'ширина', 'width'],
        'height_cm': ['height_cm', 'высота', 'height'], 'volume_liters': ['volume_liters', 'объем', 'volume'],
        'packaging_cost': ['packaging_cost', 'упаковка'], 'first_mile_cost': ['first_mile_cost', 'магистраль', 'первая миля'],
        'marketing_budget_per_unit': ['marketing_budget_per_unit', 'реклама', 'дрр'],
        'stock_depth_days': ['stock_depth_days', 'глубина запаса', 'дни запаса'], 'quantity_per_order': ['quantity_per_order', 'количество в заказе']
    }
    
    @classmethod
    def normalize_dataframe(cls, raw_df: pd.DataFrame) -> pd.DataFrame:
        if raw_df.empty: return raw_df
        df = raw_df.copy(); df.columns = [str(col).strip().lower() for col in df.columns]
        final_data = {}
        for target_col, synonyms in cls.COLUMN_MAPPING.items():
            found = False
            for syn in synonyms:
                if syn in df.columns: final_data[target_col] = df[syn]; found = True; break
            if not found: final_data[target_col] = "Кастомный SKU" if target_col == 'artikul' else ("автозапчасти" if target_col == 'category' else 0.0)
        
        norm_df = pd.DataFrame(final_data)
        num_cols = ['selling_price', 'cogs', 'daily_sales', 'weight_kg', 'length_cm', 'width_cm', 'height_cm', 'volume_liters', 'packaging_cost', 'first_mile_cost', 'marketing_budget_per_unit', 'stock_depth_days', 'quantity_per_order']
        for col in num_cols:
            if col in norm_df.columns:
                norm_df[col] = pd.to_numeric(norm_df[col].astype(str).str.replace(r'\s+|,|%|₽', '', regex=True), errors='coerce').fillna(0.0).abs()
        norm_df['artikul'] = norm_df['artikul'].astype(str).str.strip()
        norm_df['category'] = norm_df['category'].astype(str).str.strip().str.lower()
        return norm_df

    @classmethod
    def load_file(cls, file_buffer: io.BytesIO, file_name: str) -> pd.DataFrame:
        try:
            if file_name.endswith('.csv'): return pd.read_csv(file_buffer, sep=None, engine='python', encoding='utf-8')
            elif file_name.endswith(('.xls', '.xlsx')): return pd.read_excel(file_buffer)
            else: raise ValueError("Неподдерживаемый формат.")
        except UnicodeDecodeError:
            file_buffer.seek(0); return pd.read_csv(file_buffer, sep=None, engine='python', encoding='cp1251')

# ============================================================================
# БЛОК 6: STREAMLIT ИНТЕРФЕЙС
# ============================================================================
def main():
    st.set_page_config(page_title=APP_NAME, page_icon="📈", layout="wide", initial_sidebar_state="expanded")
    if 'main_df' not in st.session_state: st.session_state.main_df = pd.DataFrame()
    
    st.sidebar.title("⚙️ Панель управления")
    st.sidebar.markdown(f"**{APP_NAME} v{APP_VERSION}**")
    st.sidebar.markdown("---")
    
    st.sidebar.subheader("🔐 API Доступы")
    api_key = st.sidebar.text_input("API Key Яндекс Маркета", type="password", value=st.secrets.get("MARKET_API_KEY", ""))
    business_id = st.sidebar.text_input("Business ID", value=st.secrets.get("MARKET_BUSINESS_ID", ""))
    deepseek_key = st.sidebar.text_input("API Key DeepSeek", type="password", value=st.secrets.get("DEEPSEEK_API_KEY", ""))
    use_api = st.sidebar.checkbox("🌐 Использовать API для тарифов", value=True)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("🏪 Настройки")
    scheme_label = st.sidebar.selectbox("Схема работы:", [s.value for s in YMScheme])
    tax_label = st.sidebar.selectbox("Налогообложение:", [t.label for t in TaxSystem])
    payment_freq = st.sidebar.selectbox("Частота выплат:", ["Ежемесячно (1.0%)", "Раз в 2 недели (1.3%)", "Еженедельно, 4 нед. (1.6%)", "Ежедневно (3.3%)"], index=2)
    
    page = st.sidebar.radio("Навигация:", ["📊 Дашборд", "🔥 Метрики и ABC-XYZ", "💰 Рекомендованные цены", "🗂️ Тарифы и Справочник", "💾 Импорт / Экспорт"])
    
    tm = HybridTariffManager()
    
    df_hash = hashlib.md5(str(st.session_state.main_df.to_json()).encode()).hexdigest() if not st.session_state.main_df.empty else "empty"
    t_hash = hashlib.md5(str(tm.tariffs).encode()).hexdigest()
    calc_df = run_calculations_cached(df_hash, tax_label, scheme_label, payment_freq, t_hash)
    
    if page == "📊 Дашборд":
        st.title("📊 Панель комплексной аналитики")
        if not calc_df.empty:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Всего SKU", len(calc_df))
            c2.metric("Ср. маржинальность", f"{calc_df['margin_percent'].mean():.2f}%")
            c3.metric("Общая выручка", format_number(calc_df['selling_price'].sum(), " ₽"))
            c4.metric("ОБЩАЯ ПРИБЫЛЬ", format_number(calc_df['gross_profit'].sum(), " ₽"), delta=f"{calc_df['margin_percent'].mean():.1f}% маржа")
            st.plotly_chart(px.treemap(calc_df, path=['abc_xyz'], values='gross_profit', title="Прибыль по ABC-XYZ сегментам"), use_container_width=True)
        else: st.warning("Загрузите данные на вкладке Импорт/Экспорт.")

    elif page == "🔥 Метрики и ABC-XYZ":
        st.title("🔥 Детальные метрики")
        if not calc_df.empty:
            st.dataframe(calc_df, use_container_width=True, hide_index=True)
            st.markdown("---")
            st.subheader("💾 Экспорт")
            if OPENPYXL_AVAILABLE:
                st.download_button(label="⬇️ СКАЧАТЬ ПОЛНЫЙ ОТЧЁТ С ЖИВЫМИ ФОРМУЛАМИ (.XLSX)",
                    data=UltimateExcelExporter.export_max_info(calc_df, tax_label, scheme_label, tm),
                    file_name=f"YM_UnitEconomics_Live_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                    use_container_width=True, type="primary")
            else: st.error("Установите: pip install openpyxl")

    elif page == "💰 Рекомендованные цены":
        st.title("💰 Рекомендованные цены")
        if not calc_df.empty:
            cols = ['artikul', 'category', 'selling_price', 'cogs', 'gross_profit', 'margin_percent', 'rec_price_min', 'rec_price_15', 'rec_price_25', 'profitability_status']
            st.dataframe(calc_df[[c for c in cols if c in calc_df.columns]], use_container_width=True, hide_index=True)

    elif page == "🗂️ Тарифы и Справочник":
        st.title("🗂️ Управление тарифами")
        st.info("Приоритет: API Яндекс Маркета → DeepSeek → Загруженный файл → Базовый фоллбэк (15% с предупреждением).")
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
                st.success("✅ Справочник тарифов успешно загружен и будет использован в расчетах и Excel-формулах.")
                st.rerun()
            except Exception as e: st.error(f"Ошибка загрузки: {e}")

    elif page == "💾 Импорт / Экспорт":
        st.title("💾 Центр импорта данных")
        uploaded_file = st.file_uploader("Перетащите файл с товарами (CSV/XLSX)", type=['csv', 'xlsx'])
        if uploaded_file is not None:
            try:
                raw_data = UniversalDataNormalizer.load_file(io.BytesIO(uploaded_file.getvalue()), uploaded_file.name)
                processed_df = UniversalDataNormalizer.normalize_dataframe(raw_data)
                st.session_state.main_df = processed_df
                st.success(f"✅ Данные импортированы! Позиций: {len(processed_df)}")
                st.dataframe(processed_df.head(10), use_container_width=True, hide_index=True)
            except Exception as e: st.error(f"Ошибка импорта: {str(e)}")

if __name__ == "__main__":
    main()
