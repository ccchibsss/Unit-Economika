#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
🚀 ULTIMATE UNIT ECONOMICS FOR AUTO PARTS v18.0 - FULL DYNAMIC DASHBOARD
============================================================================
Полная версия без сокращений: 50+ метрик, кэширование, безопасная нормализация,
профессиональные графики (Waterfall, Scatter) и продвинутый Excel-экспорт.
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
from enum import Enum
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List, Tuple, Optional

warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('UltimateUnitEconomics')

OPENPYXL_AVAILABLE = False
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.formatting.rule import DataBarRule, ColorScaleRule, FormulaRule
    from openpyxl.utils import get_column_letter
    OPENPYXL_AVAILABLE = True
except ImportError:
    pass

APP_VERSION = "18.0.0"
APP_NAME = "ZapStore Ultimate Unit Economics & 50+ Metrics Dashboard"

# ============================================================================
# БЛОК 0: СЛУЖЕБНЫЕ УТИЛИТЫ ТОЧНЫХ РАСЧЕТОВ
# ============================================================================
def money_round(value: float) -> float:
    """Округление денежных сумм до копеек"""
    if pd.isna(value) or np.isinf(value): 
        return 0.0
    return float(Decimal(str(value)).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP))

def percent_round(value: float) -> float:
    """Округление процентов до 2 знаков"""
    if pd.isna(value) or np.isinf(value): 
        return 0.0
    return float(Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))

def fix_double_utf8(text: str) -> str:
    """Исправление двойной кодировки UTF-8"""
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
    """Форматирование больших чисел (1M, 1K)"""
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
# БЛОК 1: КОНФИГУРАЦИИ И СТРУКТУРЫ ДАННЫХ
# ============================================================================
class TaxSystem(Enum):
    """Системы налогообложения"""
    USN_6 = ("УСН 6% (доходы)", 0.06, "revenue", 0.0)
    USN_15 = ("УСН 15% (доходы-расходы)", 0.15, "profit", 0.01)
    OSN = ("ОСН (общая с НДС 20%)", 0.20, "profit_vat", 0.0)

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

class Tariff:
    """Тариф маркетплейса"""
    def __init__(self, category: str, commission_rate: float = 0.12, min_commission: float = 35.0, 
                 magma_base: float = 30.0, magma_per_kg: float = 15.0, acquiring_fee: float = 0.018, 
                 return_fee: float = 0.05, storage_fee_per_day: float = 0.50, source: str = "Справочник"):
        self.category = category
        self.commission_rate = commission_rate
        self.min_commission = min_commission
        self.magma_base = magma_base
        self.magma_per_kg = magma_per_kg
        self.acquiring_fee = acquiring_fee
        self.return_fee = return_fee
        self.storage_fee_per_day = storage_fee_per_day
        self.source = source

# ============================================================================
# БЛОК 2: ГИБРИДНЫЙ МЕНЕДЖЕР ТАРИФОВ
# ============================================================================
class HybridTariffManager:
    """Управление тарифами по категориям"""
    DEFAULTS = {
        'default': Tariff('default', 0.12, 35, 30, 15, 0.018, 0.05, 0.50, "Локальная база"),
        'автозапчасти': Tariff('автозапчасти', 0.10, 30, 30, 15, 0.018, 0.06, 0.50, "Оферта автозапчасти"),
        'электроника': Tariff('электроника', 0.08, 30, 30, 15, 0.015, 0.04, 0.40, "Оферта электроники"),
        'одежда': Tariff('одежда', 0.15, 25, 25, 12, 0.018, 0.07, 0.60, "Оферта одежда"),
        'обувь': Tariff('обувь', 0.13, 30, 30, 15, 0.018, 0.06, 0.55, "Оферта обувь"),
        'аксессуары': Tariff('аксессуары', 0.14, 25, 25, 12, 0.018, 0.06, 0.50, "Оферта аксессуары")
    }

    def __init__(self):
        if 'tariffs' not in st.session_state:
            st.session_state.tariffs = dict(self.DEFAULTS)

    @property
    def tariffs(self):
        return st.session_state.tariffs

    def get_best_tariff(self, category_name: str) -> Tariff:
        cat_clean = str(category_name).lower().strip()
        if cat_clean in self.tariffs: 
            return self.tariffs[cat_clean]
        for k, t in self.tariffs.items():
            if k in cat_clean or cat_clean in k: 
                return t
        return self.tariffs['default']

    def to_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame([{
            'Категория': k,
            'Комиссия, %': round(t.commission_rate * 100, 2),
            'Мин. комиссия, ₽': t.min_commission,
            'Магистраль база, ₽': t.magma_base,
            'Магистраль за кг, ₽': t.magma_per_kg,
            'Эквайринг, %': round(t.acquiring_fee * 100, 2),
            'Возвраты, %': round(t.return_fee * 100, 2),
            'Хранение день, ₽': t.storage_fee_per_day,
            'Источник данных': t.source
        } for k, t in self.tariffs.items()])

# ============================================================================
# БЛОК 3: ВЕКТОРИЗОВАННЫЙ ФИНАНСОВЫЙ ДВИЖОК С 50+ МЕТРИКАМИ
# ============================================================================
@st.cache_data(ttl=3600, show_spinner=False)
def run_calculations_cached(df_hash: str, tax_label: str, tariffs_snapshot: str) -> pd.DataFrame:
    """Кэшируемая обертка для предотвращения пересчета при каждом клике в UI"""
    df = st.session_state.main_df.copy()
    if df.empty: 
        return df
    
    tax_system = TaxSystem.by_label(tax_label)
    manager = HybridTariffManager()
    
    if 'artikul' in df.columns: 
        df['artikul'] = df['artikul'].astype(str).apply(fix_double_utf8)
    if 'category' in df.columns: 
        df['category'] = df['category'].astype(str).apply(fix_double_utf8)
    
    required_cols = {
        'selling_price': 0.0, 'cogs': 0.0, 'weight_kg': 0.0, 'length_cm': 0.0,
        'width_cm': 0.0, 'height_cm': 0.0, 'packaging_cost': 0.0,
        'marketing_budget_per_unit': 0.0, 'daily_sales': 0.0, 'stock_depth_days': 0.0,
        'first_mile_cost': 0.0, 'commission': 0.0, 'return_cost': 0.0, 'warehouse_cost': 0.0
    }
    for col, default in required_cols.items():
        if col not in df.columns: 
            df[col] = default
        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(default)

    comm_rates, min_comms, magma_bases, magma_kgs = [], [], [], []
    acq_fees, ret_fees, storage_fees = [], [], []
    
    for cat in df.get('category', ['default'] * len(df)):
        t = manager.get_best_tariff(cat)
        comm_rates.append(t.commission_rate)
        min_comms.append(t.min_commission)
        magma_bases.append(t.magma_base)
        magma_kgs.append(t.magma_per_kg)
        acq_fees.append(t.acquiring_fee)
        ret_fees.append(t.return_fee)
        storage_fees.append(t.storage_fee_per_day)

    comm_rates = np.array(comm_rates)
    acq_fees = np.array(acq_fees)
    ret_fees = np.array(ret_fees)
    storage_fees = np.array(storage_fees)

    vol_weight = (df['length_cm'] * df['width_cm'] * df['height_cm']) / 5000.0
    df['billable_weight'] = np.maximum(df['weight_kg'], vol_weight)
    df['billable_weight'] = np.ceil(df['billable_weight'] * 2) / 2

    df['commission'] = np.where(df['commission'] == 0, np.maximum(df['selling_price'] * comm_rates, min_comms), df['commission'])
    df['last_mile_cost'] = np.clip(df['selling_price'] * 0.045, 60.0, 400.0)
    df['first_mile_cost'] = np.where(df['first_mile_cost'] == 0, np.array(magma_bases) + (df['billable_weight'] * np.array(magma_kgs)), df['first_mile_cost'])
    df['acquiring_cost'] = df['selling_price'] * acq_fees
    df['return_cost'] = np.where(df['return_cost'] == 0, (150.0 + (df['selling_price'] * 0.30)) * ret_fees, df['return_cost'])
    df['pick_pack_cost'] = 35.0
    df['warehouse_cost'] = np.where(df['warehouse_cost'] == 0, (df['stock_depth_days'] * df['daily_sales']) * storage_fees, df['warehouse_cost'])

    df['fixed_operational_costs'] = df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + df['packaging_cost'] + df['return_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    df['marketplace_fees'] = df['commission'] + df['last_mile_cost'] + df['acquiring_cost']
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

    tax_factor = tax_system.rate if tax_system.base == "revenue" else 0.0
    variable_fees_share = comm_rates + 0.045 + acq_fees + tax_factor
    denom = np.where((1.0 - variable_fees_share) <= 0.01, 0.5, 1.0 - variable_fees_share)
    df['rec_price_min'] = df['fixed_operational_costs'] / denom
    df['rec_price_15'] = df['fixed_operational_costs'] / (denom - 0.15)
    df['rec_price_25'] = df['fixed_operational_costs'] / (denom - 0.25)
    df['rec_price_30'] = df['fixed_operational_costs'] / (denom - 0.30)

    df['revenue_per_unit'] = df['selling_price']
    df['total_cost_per_unit'] = df['total_expenses']
    df['variable_costs'] = df['commission'] + df['last_mile_cost'] + df['acquiring_cost'] + df['return_cost']
    df['fixed_costs'] = df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + df['packaging_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
    df['contribution_margin'] = df['selling_price'] - df['variable_costs']
    df['contribution_margin_percent'] = np.where(df['selling_price'] > 0, (df['contribution_margin'] / df['selling_price']) * 100, 0.0)
    df['gross_margin_before_tax'] = df['selling_price'] - df['pre_tax_expenses']
    df['net_margin_after_tax'] = df['margin_percent']
    df['cogs_percent'] = np.where(df['selling_price'] > 0, (df['cogs'] / df['selling_price']) * 100, 0.0)
    df['logistics_percent'] = np.where(df['selling_price'] > 0, ((df['first_mile_cost'] + df['last_mile_cost']) / df['selling_price']) * 100, 0.0)
    df['commission_percent'] = np.where(df['selling_price'] > 0, (df['commission'] / df['selling_price']) * 100, 0.0)
    df['marketing_percent'] = np.where(df['selling_price'] > 0, (df['marketing_budget_per_unit'] / df['selling_price']) * 100, 0.0)
    df['total_fees_percent'] = df['cogs_percent'] + df['logistics_percent'] + df['commission_percent'] + df['marketing_percent']
    df['roi_percent'] = np.where(df['cogs'] > 0, ((df['gross_profit'] / df['cogs']) * 100), 0.0)
    df['markup_percent'] = np.where(df['cogs'] > 0, ((df['selling_price'] - df['cogs']) / df['cogs']) * 100, 0.0)
    df['break_even_units'] = np.where(df['contribution_margin'] > 0, df['fixed_costs'] / df['contribution_margin'], 0.0)
    df['safety_margin_percent'] = np.where(df['selling_price'] > df['rec_price_min'], ((df['selling_price'] - df['rec_price_min']) / df['selling_price']) * 100, 0.0)
    df['cost_per_kg'] = np.where(df['billable_weight'] > 0, (df['first_mile_cost'] + df['last_mile_cost']) / df['billable_weight'], 0.0)
    df['revenue_per_kg'] = np.where(df['billable_weight'] > 0, df['selling_price'] / df['billable_weight'], 0.0)
    df['profit_per_kg'] = np.where(df['billable_weight'] > 0, df['gross_profit'] / df['billable_weight'], 0.0)
    df['efficiency_score'] = df['margin_percent'] * 0.4 + df['roi_percent'] * 0.3 + df['safety_margin_percent'] * 0.3
    
    df['abc_category'] = np.where(df['daily_sales'] >= 10, 'A', np.where(df['daily_sales'] >= 3, 'B', 'C'))
    df['xyz_category'] = np.where(df['margin_percent'] >= 20, 'X', np.where(df['margin_percent'] >= 10, 'Y', 'Z'))
    df['abc_xyz'] = df['abc_category'] + df['xyz_category']

    money_columns = [
        'commission', 'last_mile_cost', 'first_mile_cost', 'acquiring_cost', 'return_cost', 
        'pick_pack_cost', 'warehouse_cost', 'fixed_operational_costs', 'marketplace_fees', 
        'pre_tax_expenses', 'tax_cost', 'total_expenses', 'gross_profit', 'operating_profit', 
        'rec_price_min', 'rec_price_15', 'rec_price_25', 'rec_price_30', 'variable_costs', 
        'fixed_costs', 'contribution_margin', 'gross_margin_before_tax', 'cost_per_kg', 
        'revenue_per_kg', 'profit_per_kg'
    ]
    for col in money_columns:
        if col in df.columns: 
            df[col] = df[col].apply(money_round)
            
    percent_columns = [
        'margin_percent', 'operating_margin', 'contribution_margin_percent', 'cogs_percent', 
        'logistics_percent', 'commission_percent', 'marketing_percent', 'total_fees_percent', 
        'roi_percent', 'markup_percent', 'safety_margin_percent', 'efficiency_score'
    ]
    for col in percent_columns:
        if col in df.columns: 
            df[col] = df[col].apply(percent_round)
            
    return df

# ============================================================================
# БЛОК 4: ЭКСПОРТЕРЫ С МАКСИМАЛЬНЫМ УСЛОВНЫМ ФОРМАТИРОВАНИЕМ В EXCEL
# ============================================================================
class ABCXYZExcelExporter:
    """Экспорт ABC-XYZ анализа с продвинутой визуализацией"""
    @staticmethod
    def export_abc_xyz(df: pd.DataFrame) -> bytes:
        if not OPENPYXL_AVAILABLE or df.empty: 
            return b""
        wb = Workbook()
        ws = wb.active
        ws.title = "ABC-XYZ Анализ"
        
        headers = [
            "Категория", "ABC-XYZ", "Заказы, шт.", "Текущая цена ₽", "Себестоимость ₽", 
            "Логистика ₽", "Комиссия ₽", "Расходы хранения ₽", "ЧИСТАЯ ПРИБЫЛЬ ₽"
        ]
        
        for col_idx, text in enumerate(headers, 1):
            cell = ws.cell(1, col_idx, text)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="1F4E78", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(left=Side(style="thin", color="D9D9D9"), right=Side(style="thin", color="D9D9D9"))

        for r_idx, row in enumerate(df.itertuples(index=False), 2):
            vals = [
                getattr(row, 'category', ''), 
                getattr(row, 'abc_xyz', ''), 
                getattr(row, 'daily_sales', 0), 
                getattr(row, 'selling_price', 0.0), 
                getattr(row, 'cogs', 0.0), 
                getattr(row, 'first_mile_cost', 0.0) + getattr(row, 'last_mile_cost', 0.0), 
                getattr(row, 'commission', 0.0), 
                getattr(row, 'warehouse_cost', 0.0), 
                getattr(row, 'gross_profit', 0.0)
            ]
            for c_idx, value in enumerate(vals, 1):
                cell = ws.cell(r_idx, c_idx, value)
                cell.border = Border(left=Side(style="thin", color="E0E0E0"), right=Side(style="thin", color="E0E0E0"), top=Side(style="thin", color="E0E0E0"), bottom=Side(style="thin", color="E0E0E0"))
                cell.alignment = Alignment(horizontal="right" if isinstance(value, (int, float)) else "left")

                if c_idx == 2 and isinstance(value, str):
                    if value.startswith('AX') or value.startswith('AY') or value.startswith('BX'):
                        cell.fill = PatternFill(start_color="C6EFCE", fill_type="solid")
                        cell.font = Font(bold=True, color="006100")
                    elif value.startswith('CZ') or 'Z' in value:
                        cell.fill = PatternFill(start_color="FFC7CE", fill_type="solid")
                        cell.font = Font(bold=True, color="9C0006")
                    else:
                        cell.fill = PatternFill(start_color="FFEB9C", fill_type="solid")
                        cell.font = Font(bold=True, color="9C5700")

                if c_idx == len(headers) and isinstance(value, (int, float)):
                    if value > 0:
                        cell.fill = PatternFill(start_color="E7F6FF", fill_type="solid")
                        cell.font = Font(bold=True, color="0066CC")
                    elif value < 0:
                        cell.fill = PatternFill(start_color="FFE6E6", fill_type="solid")
                        cell.font = Font(bold=True, color="CC0000")

                if c_idx == 3 and isinstance(value, (int, float)):
                    cell.number_format = '#,##0'
                elif c_idx >= 4 and isinstance(value, (int, float)):
                    cell.number_format = '#,##0.00 ₽'

        try:
            ws.conditional_formatting.add("C2:C1000", DataBarRule(start_type='min', end_type='max', color="638EC6", showValue=True))
            ws.conditional_formatting.add("I2:I1000", FormulaRule(formula=["$I2<0"], fill=PatternFill(start_color="FFCCCC", fill_type="solid")))
        except Exception:
            pass

        for col in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 22
        ws.freeze_panes = 'A2'
        
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.getvalue()

class ComprehensiveMetricsExporter:
    """Инструмент полного финансового экспорта со всеми 50+ метриками"""
    @staticmethod
    def export_full_metrics(df: pd.DataFrame) -> bytes:
        if not OPENPYXL_AVAILABLE or df.empty: 
            return b""
        wb = Workbook()
        ws = wb.active
        ws.title = "Комплексный Анализ 50+ Метрик"
        
        for col_idx, text in enumerate(df.columns, 1):
            cell = ws.cell(1, col_idx, text)
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = PatternFill(start_color="2E75B6", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")

        for i, row in enumerate(df.values, 2):
            for j, val in enumerate(row, 1):
                cell = ws.cell(i, j)
                if isinstance(val, (int, float)):
                    cell.value = val
                    cell.number_format = '#,##0.00'
                else:
                    cell.value = str(val)
                cell.border = Border(bottom=Side(style="thin", color="E0E0E0"))

        for col in range(1, len(df.columns) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 18
            
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.getvalue()

class ExcelDynamicExporter:
    """Классический финансовый экспортер с живыми формулами Excel"""
    @staticmethod
    def export(df: pd.DataFrame) -> bytes:
        if not OPENPYXL_AVAILABLE or df.empty: 
            return b""
        wb = Workbook()
        
        ws_dash = wb.active
        ws_dash.title = "📊 Сводный Отчет"
        ws_dash.cell(1, 1, "Сводный финансовый отчет Unit Economics").font = Font(size=14, bold=True)
        ws_dash.cell(3, 1, "Всего SKU в обработке:")
        ws_dash.cell(3, 2, len(df)).font = Font(bold=True)
        ws_dash.column_dimensions['A'].width = 30
        
        ws = wb.create_sheet("Расчет экономики")
        headers = [
            'Артикул', 'Категория', 'Цена продажи', 'Себестоимость', 'Комиссия маркетплейса', 
            'Магистраль', 'Последняя миля', 'Банковский эквайринг', 'Процент возвратов', 
            'Расчетный налог', 'Итого расходов', 'ЧИСТАЯ ПРИБЫЛЬ', 'МИН. ЦЕНА (0%)', 'ОПТИМАЛЬНАЯ (15%)'
        ]
        
        for col_idx, text in enumerate(headers, 1):
            cell = ws.cell(1, col_idx, text)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1F4E78", fill_type="solid")
            
        thin_side = Side(border_style="thin", color="D9D9D9")
        data_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
        
        for i, row in df.reset_index(drop=True).iterrows():
            r = i + 2
            ws.cell(r, 1, str(row.get('artikul', ''))).border = data_border
            ws.cell(r, 2, str(row.get('category', ''))).border = data_border
            ws.cell(r, 3, float(row.get('selling_price', 0))).border = data_border
            ws.cell(r, 4, float(row.get('cogs', 0))).border = data_border
            ws.cell(r, 5, float(row.get('commission', 0))).border = data_border
            ws.cell(r, 6, float(row.get('first_mile_cost', 0))).border = data_border
            ws.cell(r, 7, float(row.get('last_mile_cost', 0))).border = data_border
            ws.cell(r, 8, float(row.get('acquiring_cost', 0))).border = data_border
            ws.cell(r, 9, float(row.get('return_cost', 0))).border = data_border
            ws.cell(r, 10, float(row.get('tax_cost', 0))).border = data_border
            
            ws.cell(r, 11, f"=SUM(D{r}:J{r})").border = data_border
            ws.cell(r, 12, f"=C{r}-K{r}").border = data_border
            
            ws.cell(r, 13, float(row.get('rec_price_min', 0))).border = data_border
            ws.cell(r, 14, float(row.get('rec_price_15', 0))).border = data_border
            
            for col_idx in range(3, 15):
                if col_idx != 9: 
                    ws.cell(r, col_idx).number_format = '#,##0.00 ₽'

        for col in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 22
            
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.getvalue()

class YandexMarketApiSync:
    """Модуль выгрузки обновленных цен на маркетплейс по API"""
    @staticmethod
    def update_prices(business_id: str, api_key: str, price_data: list) -> Tuple[bool, str]:
        if not business_id or not api_key:
            return False, "Отсутствуют Business ID или API Key в настройках."
            
        url = f"https://api.partner.market.yandex.ru/v2/businesses/{business_id}/offers/update-prices"
        headers = {"Authorization": f"OAuth {api_key}", "Content-Type": "application/json"}
        offers_payload = [{"offerId": str(item['artikul']), "price": {"value": float(item['new_price']), "currencyId": "RUR"}} for item in price_data]
        
        try:
            response = requests.post(url, json={"offers": offers_payload}, headers=headers, timeout=10)
            if response.status_code == 200: 
                return True, "Цены обновлены в ЛК маркетплейса."
            return False, f"Ошибка API ({response.status_code}): {response.text}"
        except Exception as e:
            return False, f"Ошибка соединения: {str(e)}"

# ============================================================================
# БЛОК 5: АДАПТИВНЫЙ НОРМАЛИЗАТОР ДАННЫХ (ИСПРАВЛЕННЫЙ)
# ============================================================================
class UniversalDataNormalizer:
    COLUMN_MAPPING_DICTIONARY = {
        'artikul': ['artikul', 'артикул', 'код товара', 'sku', 'offer_id', 'id', 'товар'],
        'category': ['category', 'категория', 'группа', 'тип товара', 'предмет'],
        'selling_price': ['selling_price', 'цена продажи', 'цена', 'price', 'реализация', 'выручка'],
        'cogs': ['cogs', 'себестоимость', 'закупка', 'cost', 'себестоимость р.'],
        'daily_sales': ['daily_sales', 'заказы, шт.', 'продажи, шт.', 'quantity', 'sales_count', 'заказы'],
        'first_mile_cost': ['first_mile_cost', 'магистраль', 'логистика', 'доставка'],
        'commission': ['commission', 'комиссия', 'marketplace_fee'],
        'return_cost': ['return_cost', 'возвраты', 'возвраты ₽'],
        'marketing_budget_per_unit': ['marketing_budget_per_unit', 'рекламные затраты', 'реклама', 'marketing'],
        'warehouse_cost': ['warehouse_cost', 'стоимость хранения', 'хранение', 'storage']
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
        numeric_cols = ['selling_price', 'cogs', 'daily_sales', 'first_mile_cost', 'commission', 'return_cost', 'marketing_budget_per_unit', 'warehouse_cost']
        
        for col in numeric_cols:
            if col in normalized_df.columns:
                normalized_df[col] = normalized_df[col].astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.')
                normalized_df[col] = pd.to_numeric(normalized_df[col], errors='coerce').fillna(0.0)
                
                if col in ['cogs', 'first_mile_cost', 'commission', 'return_cost', 'marketing_budget_per_unit', 'warehouse_cost']:
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
    labels = ["Цена продажи", "Себестоимость", "Логистика (1+ посл. миля)", "Комиссия МП", "Прочие расходы", "Налог", "Чистая прибыль"]
    
    price = avg_row.get('selling_price', 0)
    cogs = -avg_row.get('cogs', 0)
    logistics = -(avg_row.get('first_mile_cost', 0) + avg_row.get('last_mile_cost', 0))
    commission = -avg_row.get('commission', 0)
    other = -(avg_row.get('acquiring_cost', 0) + avg_row.get('return_cost', 0) + avg_row.get('pick_pack_cost', 0) + avg_row.get('warehouse_cost', 0) + avg_row.get('marketing_budget_per_unit', 0))
    tax = -avg_row.get('tax_cost', 0)
    profit = avg_row.get('gross_profit', 0)
    
    values = [price, cogs, logistics, commission, other, tax, profit]
    measure = ["absolute", "relative", "relative", "relative", "relative", "relative", "total"]
    
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
    fig.update_layout(title="Структура цены среднего товара (Waterfall)", showlegend=False)
    return fig

def render_margin_scatter(df: pd.DataFrame) -> go.Figure:
    if df.empty: 
        return go.Figure()
    
    fig = px.scatter(
        df, x="daily_sales", y="margin_percent", 
        size="gross_profit", color="abc_xyz",
        hover_data=["artikul", "selling_price", "gross_profit"],
        title="Матрица эффективности: Маржинальность (%) vs Продажи (шт/день)",
        labels={"daily_sales": "Продажи (шт/день)", "margin_percent": "Маржинальность (%)"},
        color_discrete_map={
            'AX': '#006100', 'AY': '#2ca02c', 'BX': '#1f77b4', 'BY': '#17becf',
            'CX': '#ff7f0e', 'CY': '#bcbd22', 'CZ': '#d62728', 'AZ': '#e377c2'
        }
    )
    fig.add_hline(y=0, line_dash="dash", line_color="red", annotation_text="Точка безубыточности")
    fig.update_layout(height=500)
    return fig

# ============================================================================
# БЛОК 7: STREAMLIT ИНТЕРФЕЙС И ОПЕРАЦИОННЫЕ ЭКРАНЫ
# ============================================================================
def main():
    st.set_page_config(page_title=APP_NAME, page_icon="📈", layout="wide", initial_sidebar_state="expanded")
    
    if 'main_df' not in st.session_state:
        st.session_state.main_df = pd.DataFrame(columns=['artikul', 'category', 'selling_price', 'cogs', 'daily_sales'])

    st.sidebar.title("⚙️ Панель управления")
    st.sidebar.markdown(f"**{APP_NAME}**<br>Версия: {APP_VERSION} © 2026", unsafe_allow_html=True)
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔐 Доступы API")
    st.sidebar.info("Ключи не хранятся в коде. Введите их здесь или используйте st.secrets.")
    api_key = st.sidebar.text_input("API Key", type="password", value=st.secrets.get("MARKET_API_KEY", ""))
    business_id = st.sidebar.text_input("Business ID", value=st.secrets.get("MARKET_BUSINESS_ID", ""))
    
    st.sidebar.markdown("---")
    tax_label = st.sidebar.selectbox("Налогообложение компании:", [t.label for t in TaxSystem])
    
    page = st.sidebar.radio("Выберите рабочий экран:", [
        "📊 Сводный Дашборд", "🔥 ABC-XYZ и 50+ Метрик", 
        "📝 Калькулятор экономики", "🗂️ Управление категориями", 
        "💾 Импорт / Экспорт данных", "📡 Синхронизация API"
    ])
    
    df_hash = hashlib.md5(str(st.session_state.main_df.to_json()).encode()).hexdigest() if not st.session_state.main_df.empty else "empty"
    tariffs_snapshot = str(st.session_state.get('tariffs', {}))
    
    calculated_df = run_calculations_cached(df_hash, tax_label, tariffs_snapshot)

    if page == "📊 Сводный Дашборд":
        st.title("📊 Панель комплексной аналитики")
        if not calculated_df.empty:
            c1, c2, c3, c4 = st.columns(4)
            c1.metric("Всего SKU в матрице", f"{len(calculated_df)}")
            c2.metric("Ср. маржинальность портфеля", f"{calculated_df['margin_percent'].mean():.2f}%")
            c3.metric("Общая выручка", format_number(calculated_df['selling_price'].sum(), " ₽"))
            c4.metric("ОБЩАЯ ПРИБЫЛЬ", format_number(calculated_df['gross_profit'].sum(), " ₽"))
            
            st.markdown("---")
            col1, col2 = st.columns(2)
            with col1:
                st.plotly_chart(render_waterfall_chart(calculated_df), use_container_width=True)
            with col2:
                avg_costs = calculated_df[['cogs', 'first_mile_cost', 'last_mile_cost', 'commission', 'tax_cost']].mean().reset_index()
                avg_costs.columns = ['Статья', 'Сумма']
                fig_tree = px.treemap(avg_costs, path=['Статья'], values='Сумма', 
                                      title="Средняя структура расходов на единицу",
                                      color='Сумма', color_continuous_scale='Reds')
                st.plotly_chart(fig_tree, use_container_width=True)
                
            st.subheader("📋 Ключевые метрики SKU")
            st.dataframe(calculated_df[['artikul', 'category', 'selling_price', 'cogs', 'gross_profit', 'margin_percent']], use_container_width=True, hide_index=True)
        else:
            st.warning("Товарная матрица пуста. Пожалуйста, загрузите свой рабочий файл на вкладке '💾 Импорт / Экспорт данных'.")

    elif page == "🔥 ABC-XYZ и 50+ Метрик":
        st.title("🔥 ABC-XYZ Матрица и Метрики Ваших Товаров")
        if not calculated_df.empty:
            st.plotly_chart(render_margin_scatter(calculated_df), use_container_width=True)
            
            st.subheader("50+ Метрик эффективности (полный детализированный анализ)")
            st.dataframe(calculated_df, use_container_width=True, hide_index=True)
            
            st.markdown("---")
            st.subheader("💾 Экспорт визуализированных Excel-отчетов")
            if OPENPYXL_AVAILABLE:
                c_exp1, c_exp2, c_exp3 = st.columns(3)
                with c_exp1:
                    abc_bytes = ABCXYZExcelExporter.export_abc_xyz(calculated_df)
                    st.download_button("⬇️ ABC-XYZ ВИЗУАЛИЗИРОВАННЫЙ (.XLSX)", data=abc_bytes, file_name=f"ABC_XYZ_Visual_{datetime.now().strftime('%d_%m_%Y')}.xlsx", use_container_width=True)
                with c_exp2:
                    metrics_excel = ComprehensiveMetricsExporter.export_full_metrics(calculated_df)
                    st.download_button("⬇️ ВСЕ 50+ МЕТРИК (.XLSX)", data=metrics_excel, file_name=f"Full_Metrics_{datetime.now().strftime('%d_%m_%Y')}.xlsx", use_container_width=True)
                with c_exp3:
                    live_excel = ExcelDynamicExporter.export(calculated_df)
                    st.download_button("⬇️ ЖИВЫЕ ФОРМУЛЫ EXCEL (.XLSX)", data=live_excel, file_name=f"Live_Formulas_{datetime.now().strftime('%d_%m_%Y')}.xlsx", use_container_width=True)
        else:
            st.info("Добавьте ваши данные для формирования матрицы условного форматирования.")

    elif page == "📝 Калькулятор экономики":
        st.title("📝 Симулятор товарной матрицы")
        if not st.session_state.main_df.empty:
            edited_df = st.data_editor(st.session_state.main_df, num_rows="dynamic", use_container_width=True, key="product_editor", hide_index=True)
            if not edited_df.equals(st.session_state.main_df):
                st.session_state.main_df = edited_df
                st.rerun()
        else:
            st.info("Матрица пуста. Загрузите файл или введите данные вручную в редактор.")
            dummy_df = pd.DataFrame([{ 'artikul': 'TEST-SKU', 'category': 'автозапчасти', 'selling_price': 1000.0, 'cogs': 500.0, 'daily_sales': 1, 'weight_kg': 1.0 }])
            st.session_state.main_df = st.data_editor(dummy_df, num_rows="dynamic", use_container_width=True, hide_index=True)

    elif page == "🗂️ Управление категориями":
        st.title("🗂️ Индивидуальные тарифы ваших категорий")
        tm = HybridTariffManager()
        with st.form("add_category_form"):
            col1, col2 = st.columns(2)
            with col1:
                new_cat = st.text_input("Название вашей категории товара:")
            with col2:
                new_comm = st.number_input("Процент комиссии маркетплейса (%)", value=12.0) / 100
            submitted = st.form_submit_button("➕ Сохранить тарифы категории")
            
            if submitted and new_cat:
                tm.tariffs[new_cat.lower().strip()] = Tariff(category=new_cat.lower().strip(), commission_rate=new_comm, source="Пользовательская база")
                st.success(f"Категория {new_cat} добавлена в расчетное ядро!")
                st.rerun()
        st.dataframe(tm.to_dataframe(), use_container_width=True, hide_index=True)

    elif page == "💾 Импорт / Экспорт данных":
        st.title("Централизованный импорт и экспорт ваших данных")
        st.markdown("### 📋 Загрузите ваш собственный файл. Поиск соответствий названий столбцов и исправление знаков произойдет автоматически.")
        
        uploaded_file = st.file_uploader("Перетащите файл CSV, XLSX или JSON с вашими данными сюда", type=['csv', 'xlsx', 'json'])
        if uploaded_file is not None:
            try:
                bytes_data = uploaded_file.getvalue()
                raw_data = UniversalDataNormalizer.load_file_dynamically(io.BytesIO(bytes_data), uploaded_file.name)
                processed_df = UniversalDataNormalizer.normalize_dataframe(raw_data)
                
                if not processed_df.empty:
                    st.session_state.main_df = processed_df
                    st.success(f"✅ Ваши данные успешно импортированы! Позиций в матрице: {len(processed_df)}")
                    st.dataframe(processed_df.head(5), use_container_width=True, hide_index=True)
                else:
                    st.error("В файле отсутствуют данные для считывания.")
            except Exception as e:
                st.error(f"Не удалось распознать структуру файла: {str(e)}")

    elif page == "📡 Синхронизация API":
        st.title("📡 Управление API-шлюзом ZapStore")
        selected_strategy = st.selectbox("Выберите стратегию цены для отправки в магазин:", ["Текущая цена", "Минимальная цена (Безубыточность)", "Оптимальная цена (15% маржа)"])
        
        if st.button("🚀 ВЫГРУЗИТЬ ОБНОВЛЕННЫЕ ЦЕНЫ НА МАРКЕТПЛЕЙС", type="primary", use_container_width=True) and not calculated_df.empty:
            price_data_to_send = []
            for _, row in calculated_df.iterrows():
                target_price = row.get('selling_price', 0) if selected_strategy == "Текущая цена" else (row.get('rec_price_min', 0) if selected_strategy == "Минимальная цена (Безубыточность)" else row.get('rec_price_15', 0))
                price_data_to_send.append({'artikul': row.get('artikul', ''), 'new_price': target_price})
            
            with st.spinner("Отправка пакета данных на сервер..."):
                success, msg = YandexMarketApiSync.update_prices(business_id, api_key, price_data_to_send)
                if success:
                    st.success(f"✅ Успешно! {msg}")
                else:
                    st.error(f"❌ Сбой выгрузки: {msg}")

if __name__ == "__main__":
    main()
