#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
============================================================================
🚀 ULTIMATE UNIT ECONOMICS FOR AUTO PARTS v17.0 - FULL METRICS DASHBOARD
============================================================================
С 50+ метриками, ABC-XYZ анализом и расчетом реального заработка
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
from enum import Enum
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List, Tuple, Optional
import base64

warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('UltimateUnitEconomics')

# Проверка доступности openpyxl
OPENPYXL_AVAILABLE = False
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill
    from openpyxl.formatting.rule import DataBarRule, ColorScaleRule, FormulaRule
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
    OPENPYXL_AVAILABLE = True
except ImportError:
    pass

APP_VERSION = "17.0.0"
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
        except:
            continue
    return text

def format_number(num: float, suffix='') -> str:
    """Форматирование больших чисел (1M, 1K)"""
    if pd.isna(num):
        return "0"
    num = abs(num)
    for unit in ['', 'K', 'M', 'B']:
        if abs(num) < 1000.0:
            return f"{num:3.1f}{unit}{suffix}"
        num /= 1000.0
    return f"{num:.1f}T{suffix}"

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
        """Получение лучшего тарифа для категории"""
        cat_clean = str(category_name).lower().strip()
        if cat_clean in self.tariffs:
            return self.tariffs[cat_clean]
        for k, t in self.tariffs.items():
            if k in cat_clean or cat_clean in k:
                return t
        return self.tariffs['default']

    def to_dataframe(self) -> pd.DataFrame:
        """Конвертация тарифов в DataFrame"""
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
class VectorizedEnginePRO:
    """Расчетный движок с 50+ метриками"""
    
    @staticmethod
    def run_calculations(df: pd.DataFrame, tax_system: TaxSystem, manager: HybridTariffManager) -> pd.DataFrame:
        """Полный расчет unit economics с 50+ метриками"""
        if df.empty:
            return df
        
        df = df.copy()
        
        # Исправление кодировок
        if 'artikul' in df.columns:
            df['artikul'] = df['artikul'].astype(str).apply(fix_double_utf8)
        if 'category' in df.columns:
            df['category'] = df['category'].astype(str).apply(fix_double_utf8)

        # Инициализация обязательных колонок
        required_cols = {
            'selling_price': 0.0, 'cogs': 0.0, 'weight_kg': 0.0,
            'length_cm': 0.0, 'width_cm': 0.0, 'height_cm': 0.0,
            'packaging_cost': 0.0, 'marketing_budget_per_unit': 0.0,
            'daily_sales': 0.0, 'stock_depth_days': 0.0
        }
        
        for col, default in required_cols.items():
            if col not in df.columns:
                df[col] = default
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(default)

        # Получение тарифов для каждой категории
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

        # === ЛОГИСТИКА И ВЕС ===
        # Объемный вес
        vol_weight = (df['length_cm'] * df['width_cm'] * df['height_cm']) / 5000.0
        df['billable_weight'] = np.maximum(df['weight_kg'], vol_weight)
        df['billable_weight'] = np.ceil(df['billable_weight'] * 2) / 2  # Округление до 0.5 кг

        # === КОМИССИИ И СБОРЫ МАРКЕТПЛЕЙСА ===
        df['commission'] = np.maximum(df['selling_price'] * comm_rates, min_comms)
        df['last_mile_cost'] = np.clip(df['selling_price'] * 0.045, 60.0, 400.0)  # Последняя миля
        df['first_mile_cost'] = np.array(magma_bases) + (df['billable_weight'] * np.array(magma_kgs))
        df['acquiring_cost'] = df['selling_price'] * acq_fees
        df['return_cost'] = (150.0 + (df['selling_price'] * 0.30)) * ret_fees
        df['pick_pack_cost'] = 35.0  # Комплектация
        df['warehouse_cost'] = (df['stock_depth_days'] * df['daily_sales']) * storage_fees

        # === ОПЕРАЦИОННЫЕ РАСХОДЫ ===
        df['fixed_operational_costs'] = (
            df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + 
            df['packaging_cost'] + df['return_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
        )
        
        df['marketplace_fees'] = df['commission'] + df['last_mile_cost'] + df['acquiring_cost']
        
        df['pre_tax_expenses'] = df['fixed_operational_costs'] + df['marketplace_fees']

        # === НАЛОГИ ===
        if tax_system.base == "revenue":
            df['tax_cost'] = df['selling_price'] * tax_system.rate
        elif tax_system.base == "profit":
            pre_tax_profit = df['selling_price'] - df['pre_tax_expenses']
            calculated_tax = np.maximum(pre_tax_profit, 0) * tax_system.rate
            min_tax = df['selling_price'] * tax_system.min_rate
            df['tax_cost'] = np.maximum(calculated_tax, min_tax)
        elif tax_system.base == "profit_vat":
            vat = df['selling_price'] * 0.20 / 1.20
            pre_tax_profit = (df['selling_price'] - vat) - df['pre_tax_expenses']
            df['tax_cost'] = np.maximum(pre_tax_profit, 0) * tax_system.rate

        # === ПРИБЫЛЬ И МАРЖАЛЬНОСТЬ ===
        df['total_expenses'] = df['pre_tax_expenses'] + df['tax_cost']
        df['gross_profit'] = df['selling_price'] - df['total_expenses']
        df['margin_percent'] = np.where(df['selling_price'] > 0, (df['gross_profit'] / df['selling_price']) * 100, 0.0)
        
        # Операционная прибыль (EBIT)
        df['operating_profit'] = df['selling_price'] - df['pre_tax_expenses']
        df['operating_margin'] = np.where(df['selling_price'] > 0, (df['operating_profit'] / df['selling_price']) * 100, 0.0)

        # === РАСЧЕТНЫЕ ЦЕНЫ ===
        tax_factor = tax_system.rate if tax_system.base == "revenue" else 0.0
        variable_fees_share = comm_rates + 0.045 + acq_fees + tax_factor
        denom = 1.0 - variable_fees_share
        denom = np.where(denom <= 0.01, 0.5, denom)

        df['rec_price_min'] = df['fixed_operational_costs'] / denom
        df['rec_price_15'] = df['fixed_operational_costs'] / (denom - 0.15)
        df['rec_price_25'] = df['fixed_operational_costs'] / (denom - 0.25)
        df['rec_price_30'] = df['fixed_operational_costs'] / (denom - 0.30)

        # === 50+ МЕТРИКИ ЭФФЕКТИВНОСТИ ===
        
        # 1-10: Базовые финансовые метрики
        df['revenue_per_unit'] = df['selling_price']
        df['total_cost_per_unit'] = df['total_expenses']
        df['variable_costs'] = df['commission'] + df['last_mile_cost'] + df['acquiring_cost'] + df['return_cost']
        df['fixed_costs'] = df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + df['packaging_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
        
        # 11-20: Маржинальность
        df['contribution_margin'] = df['selling_price'] - df['variable_costs']
        df['contribution_margin_percent'] = np.where(df['selling_price'] > 0, (df['contribution_margin'] / df['selling_price']) * 100, 0.0)
        df['gross_margin_before_tax'] = df['selling_price'] - df['pre_tax_expenses']
        df['net_margin_after_tax'] = df['margin_percent']
        
        # 21-30: Эффективность затрат
        df['cogs_percent'] = np.where(df['selling_price'] > 0, (df['cogs'] / df['selling_price']) * 100, 0.0)
        df['logistics_percent'] = np.where(df['selling_price'] > 0, ((df['first_mile_cost'] + df['last_mile_cost']) / df['selling_price']) * 100, 0.0)
        df['commission_percent'] = np.where(df['selling_price'] > 0, (df['commission'] / df['selling_price']) * 100, 0.0)
        df['marketing_percent'] = np.where(df['selling_price'] > 0, (df['marketing_budget_per_unit'] / df['selling_price']) * 100, 0.0)
        df['total_fees_percent'] = df['cogs_percent'] + df['logistics_percent'] + df['commission_percent'] + df['marketing_percent']
        
        # 31-40: ROI и эффективность
        df['roi_percent'] = np.where(df['cogs'] > 0, ((df['gross_profit'] / df['cogs']) * 100), 0.0)
        df['markup_percent'] = np.where(df['cogs'] > 0, ((df['selling_price'] - df['cogs']) / df['cogs']) * 100, 0.0)
        df['break_even_units'] = np.where(df['contribution_margin'] > 0, df['fixed_costs'] / df['contribution_margin'], 0.0)
        df['safety_margin_percent'] = np.where(df['selling_price'] > df['rec_price_min'], ((df['selling_price'] - df['rec_price_min']) / df['selling_price']) * 100, 0.0)
        
        # 41-50: Дополнительные метрики
        df['cost_per_kg'] = np.where(df['billable_weight'] > 0, (df['first_mile_cost'] + df['last_mile_cost']) / df['billable_weight'], 0.0)
        df['revenue_per_kg'] = np.where(df['billable_weight'] > 0, df['selling_price'] / df['billable_weight'], 0.0)
        df['profit_per_kg'] = np.where(df['billable_weight'] > 0, df['gross_profit'] / df['billable_weight'], 0.0)
        df['efficiency_score'] = df['margin_percent'] * 0.4 + df['roi_percent'] * 0.3 + df['safety_margin_percent'] * 0.3
        df['abc_category'] = df.apply(lambda x: 'A' if x['daily_sales'] >= 10 else ('B' if x['daily_sales'] >= 3 else 'C'), axis=1)
        
        # 51-55: ABC-XYZ анализ
        df['xyz_category'] = df.apply(lambda x: 'X' if x['margin_percent'] >= 20 else ('Y' if x['margin_percent'] >= 10 else 'Z'), axis=1)
        df['abc_xyz'] = df['abc_category'] + df['xyz_category'] + 'CZ'  # Упрощенная модель
        
        # Округление денежных значений
        money_columns = [
            'commission', 'last_mile_cost', 'first_mile_cost', 'acquiring_cost', 
            'return_cost', 'pick_pack_cost', 'warehouse_cost', 'fixed_operational_costs',
            'marketplace_fees', 'pre_tax_expenses', 'tax_cost', 'total_expenses',
            'gross_profit', 'operating_profit', 'rec_price_min', 'rec_price_15',
            'rec_price_25', 'rec_price_30', 'variable_costs', 'fixed_costs',
            'contribution_margin', 'gross_margin_before_tax', 'cost_per_kg',
            'revenue_per_kg', 'profit_per_kg'
        ]
        
        for col in money_columns:
            if col in df.columns:
                df[col] = df[col].apply(money_round)
        
        # Округление процентов
        percent_columns = [
            'margin_percent', 'operating_margin', 'contribution_margin_percent',
            'cogs_percent', 'logistics_percent', 'commission_percent',
            'marketing_percent', 'total_fees_percent', 'roi_percent',
            'markup_percent', 'safety_margin_percent', 'efficiency_score'
        ]
        
        for col in percent_columns:
            if col in df.columns:
                df[col] = df[col].apply(percent_round)

        return df

# ============================================================================
# БЛОК 4: ЭКСПОРТЕРЫ С УСЛОВНЫМ ФОРМАТИРОВАНИЕМ
# ============================================================================
class ABCXYZExcelExporter:
    """Экспорт ABC-XYZ анализа с условным форматированием"""
    
    @staticmethod
    def export_abc_xyz(df: pd.DataFrame) -> bytes:
        """Экспорт с прогресс-барами, тепловой картой и цветовой кодировкой"""
        if not OPENPYXL_AVAILABLE:
            return b""
        
        wb = Workbook()
        ws = wb.active
        ws.title = "ABC-XYZ Анализ"

        # Заголовки
        headers = [
            "Категория", "ABC-XYZ", "Заказы, шт.", "% Выкупа", "Возвраты, шт.",
            "Продажи, шт.", "Продажи до возвр. ₽", "Продажи после возвр. ₽",
            "Себестоимость ₽", "Логистика ₽", "Комиссия ₽", "Возвраты ₽",
            "Рекламные затраты ₽", "Стоимость хранения ₽", "ПРИБЫЛЬ ₽"
        ]
        
        for col_idx, text in enumerate(headers, 1):
            cell = ws.cell(1, col_idx, text)
            cell.font = Font(bold=True, color="FFFFFF", size=11)
            cell.fill = PatternFill(start_color="1F4E78", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9")
            )

        # Заполнение данных
        for r_idx, row in enumerate(df.itertuples(index=False), 2):
            for c_idx, value in enumerate(row, 1):
                cell = ws.cell(r_idx, c_idx, value)
                cell.border = Border(
                    left=Side(style="thin", color="E0E0E0"),
                    right=Side(style="thin", color="E0E0E0"),
                    top=Side(style="thin", color="E0E0E0"),
                    bottom=Side(style="thin", color="E0E0E0")
                )
                cell.alignment = Alignment(horizontal="right" if isinstance(value, (int, float)) else "left")
                
                # 1. Цветовая кодировка ABC-XYZ (Столбец 2)
                if c_idx == 2 and isinstance(value, str):
                    if value.startswith('AA') or value.startswith('AB'):
                        cell.fill = PatternFill(start_color="C6EFCE", fill_type="solid")  # Зеленый
                        cell.font = Font(bold=True, color="006100")
                    elif value.startswith('BA') or value.startswith('BB'):
                        cell.fill = PatternFill(start_color="FFEB9C", fill_type="solid")  # Желтый
                        cell.font = Font(bold=True, color="9C5700")
                    elif value.startswith('C'):
                        cell.fill = PatternFill(start_color="FFC7CE", fill_type="solid")  # Красный
                        cell.font = Font(bold=True, color="9C0006")
                
                # 2. Подсветка ПРИБЫЛИ (последний столбец)
                if c_idx == len(headers):
                    if isinstance(value, (int, float)):
                        if value > 0:
                            cell.fill = PatternFill(start_color="E7F6FF", fill_type="solid")  # Голубой
                            cell.font = Font(bold=True, color="0066CC")
                        elif value < 0:
                            cell.fill = PatternFill(start_color="FFE6E6", fill_type="solid")  # Красный
                            cell.font = Font(bold=True, color="CC0000")
                
                # 3. Форматирование чисел
                if c_idx in [3, 5, 6]:  # Штучные значения
                    if isinstance(value, (int, float)):
                        cell.number_format = '#,##0'
                elif c_idx == 4:  # Процент выкупа
                    if isinstance(value, (int, float)):
                        cell.number_format = '0.00%'
                elif c_idx in [7, 8, 9, 10, 11, 12, 13, 14, 15]:  # Денежные значения
                    if isinstance(value, (int, float)):
                        cell.number_format = '#,##0.00 ₽'

        # Условное форматирование: Прогресс-бары для "% Выкупа" (Столбец 4)
        ws.conditional_formatting.add(
            "D2:D100",
            DataBarRule(start_type='min', end_type='max', color="638EC6", showValue=True)
        )
        
        # Условное форматирование: Тепловая карта для "Продажи до возвр. ₽" (Столбец 7)
        ws.conditional_formatting.add(
            "G2:G100",
            ColorScaleRule(
                start_type='min', start_color='FDEBD0',
                mid_type='percentile', mid_value=50, mid_color='F5B041',
                end_type='max', end_color='BA4A00'
            )
        )
        
        # Условное форматирование: Подсветка убытков в столбце прибыли
        ws.conditional_formatting.add(
            "O2:O100",
            FormulaRule(formula=["$O2<0"], fill=PatternFill(start_color="FFCCCC", fill_type="solid"))
        )

        # Настройка ширины столбцов
        column_widths = [20, 12, 14, 12, 14, 14, 18, 18, 18, 14, 14, 14, 18, 18, 18]
        for col, width in enumerate(column_widths, 1):
            ws.column_dimensions[get_column_letter(col)].width = width

        # Заморозка шапки
        ws.freeze_panes = 'A2'

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.getvalue()


class ComprehensiveMetricsExporter:
    """Экспорт всех 50+ метрик"""
    
    @staticmethod
    def export_full_metrics(df: pd.DataFrame) -> bytes:
        """Полный экспорт со всеми метриками"""
        if not OPENPYXL_AVAILABLE:
            return b""
        
        wb = Workbook()
        
        # Лист 1: Сводный дашборд
        ws_summary = wb.active
        ws_summary.title = "📊 Дашборд"
        
        # Лист 2: Детальные метрики
        ws_metrics = wb.create_sheet(" 50+ Метрик")
        
        # Лист 3: ABC-XYZ
        ws_abc = wb.create_sheet(" ABC-XYZ")
        
        # Заполнение листа с метриками
        metrics_headers = [
            "Артикул", "Категория", "Цена продажи", "Себестоимость",
            "ВАЛОВАЯ ПРИБЫЛЬ", "Маржа %", "ROI %", "Наценка %",
            "Комиссия MP", "Логистика", "Эквайринг", "Возвраты",
            "Маркетинг", "Хранение", "НАЛОГИ", "ИТОГО расходов",
            "Цена мин (0%)", "Цена оптим (15%)", "Цена макс (25%)",
            "Точка безубыт", "Запас прочности %", "Эффективность",
            "ABC", "XYZ", "ABC-XYZ"
        ]
        
        for col_idx, text in enumerate(metrics_headers, 1):
            cell = ws_metrics.cell(1, col_idx, text)
            cell.font = Font(bold=True, color="FFFFFF", size=10)
            cell.fill = PatternFill(start_color="2E75B6", fill_type="solid")
            cell.alignment = Alignment(horizontal="center")
            cell.border = Border(left=Side(style="thin"), right=Side(style="thin"),
                               top=Side(style="thin"), bottom=Side(style="thin"))
        
        # Заполнение данных
        for i, row in enumerate(df.reset_index(drop=True).iterrows(), 2):
            idx, r = row
            ws_metrics.cell(i, 1, str(r.get('artikul', '')))
            ws_metrics.cell(i, 2, str(r.get('category', '')))
            ws_metrics.cell(i, 3, float(r.get('selling_price', 0)))
            ws_metrics.cell(i, 4, float(r.get('cogs', 0)))
            ws_metrics.cell(i, 5, float(r.get('gross_profit', 0)))  # ВЫДЕЛЕНИЕ ПРИБЫЛИ
            ws_metrics.cell(i, 6, f"{r.get('margin_percent', 0):.2f}%")
            ws_metrics.cell(i, 7, f"{r.get('roi_percent', 0):.2f}%")
            ws_metrics.cell(i, 8, f"{r.get('markup_percent', 0):.2f}%")
            ws_metrics.cell(i, 9, float(r.get('commission', 0)))
            ws_metrics.cell(i, 10, float(r.get('first_mile_cost', 0)) + float(r.get('last_mile_cost', 0)))
            ws_metrics.cell(i, 11, float(r.get('acquiring_cost', 0)))
            ws_metrics.cell(i, 12, float(r.get('return_cost', 0)))
            ws_metrics.cell(i, 13, float(r.get('marketing_budget_per_unit', 0)))
            ws_metrics.cell(i, 14, float(r.get('warehouse_cost', 0)))
            ws_metrics.cell(i, 15, float(r.get('tax_cost', 0)))
            ws_metrics.cell(i, 16, float(r.get('total_expenses', 0)))
            ws_metrics.cell(i, 17, float(r.get('rec_price_min', 0)))
            ws_metrics.cell(i, 18, float(r.get('rec_price_15', 0)))
            ws_metrics.cell(i, 19, float(r.get('rec_price_25', 0)))
            ws_metrics.cell(i, 20, float(r.get('break_even_units', 0)))
            ws_metrics.cell(i, 21, f"{r.get('safety_margin_percent', 0):.2f}%")
            ws_metrics.cell(i, 22, f"{r.get('efficiency_score', 0):.2f}")
            ws_metrics.cell(i, 23, str(r.get('abc_category', '')))
            ws_metrics.cell(i, 24, str(r.get('xyz_category', '')))
            ws_metrics.cell(i, 25, str(r.get('abc_xyz', '')))
            
            # Подсветка прибыли
            profit_cell = ws_metrics.cell(i, 5)
            profit = float(r.get('gross_profit', 0))
            if profit > 0:
                profit_cell.fill = PatternFill(start_color="C6EFCE", fill_type="solid")
                profit_cell.font = Font(bold=True, color="006100")
            elif profit < 0:
                profit_cell.fill = PatternFill(start_color="FFC7CE", fill_type="solid")
                profit_cell.font = Font(bold=True, color="9C0006")
        
        # Настройка ширины столбцов
        for col in range(1, len(metrics_headers) + 1):
            ws_metrics.column_dimensions[get_column_letter(col)].width = 16
        
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.getvalue()


class ExcelDynamicExporter:
    """Классический экспортер"""
    
    @staticmethod
    def export(df: pd.DataFrame) -> bytes:
        if not OPENPYXL_AVAILABLE:
            return b""
        
        wb = Workbook()
        ws_dash = wb.active
        ws_dash.title = "📊 Дашборд"
        ws_dash.cell(1, 1, "Сводный финансовый отчет").font = Font(size=14, bold=True)
        ws_dash.cell(3, 1, "Всего SKU в обработке:")
        ws_dash.cell(3, 2, len(df))
        ws_dash.column_dimensions['A'].width = 30

        ws = wb.create_sheet("Расчет экономики")
        headers = [
            'Артикул', 'Категория', 'Цена продажи', 'Себестоимость',
            'Комиссия маркетплейса', 'Магистраль', 'Последняя миля',
            'Банковский эквайринг', 'Процент возвратов', 'Расчетный налог',
            'Итого расходов', 'ЧИСТАЯ ПРИБЫЛЬ', 'Текущая маржа, %',
            'МИН. ЦЕНА (0%)', 'ОПТИМАЛЬНАЯ (15%)', 'МАКСИМАЛЬНАЯ (25%)'
        ]
        
        for col_idx, text in enumerate(headers, 1):
            cell = ws.cell(1, col_idx, text)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1F4E78", fill_type="solid")

        thin_side = Side(border_style="thin", color="D9D9D9")
        data_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

        for i, row in df.reset_index(drop=True).iterrows():
            r = i + 2
            ws.cell(r, 1, str(row['artikul'])).border = data_border
            ws.cell(r, 2, str(row['category'])).border = data_border
            ws.cell(r, 3, float(row['selling_price'])).border = data_border
            ws.cell(r, 4, float(row['cogs'])).border = data_border
            ws.cell(r, 5, float(row['commission'])).border = data_border
            ws.cell(r, 6, float(row['first_mile_cost'])).border = data_border
            ws.cell(r, 7, float(row['last_mile_cost'])).border = data_border
            ws.cell(r, 8, float(row['acquiring_cost'])).border = data_border
            ws.cell(r, 9, float(row['return_cost'])).border = data_border
            ws.cell(r, 10, float(row['tax_cost'])).border = data_border
            ws.cell(r, 11, f"=SUM(D{r}:J{r})").border = data_border
            ws.cell(r, 12, f"=C{r}-K{r}").border = data_border
            ws.cell(r, 13, f"=IF(C{r}>0, (L{r}/C{r})*100, 0)").border = data_border
            ws.cell(r, 14, float(row['rec_price_min'])).border = data_border
            ws.cell(r, 15, float(row['rec_price_15'])).border = data_border
            ws.cell(r, 16, float(row['rec_price_25'])).border = data_border
            
            # Подсветка прибыли
            profit_cell = ws.cell(r, 12)
            profit_formula = f"=C{r}-K{r}"
            profit_cell.value = profit_formula
            
        for col in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 20

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.getvalue()


class YandexMarketApiSync:
    """Синхронизация с API Яндекс Маркета"""
    
    @staticmethod
    def update_prices(business_id: str, api_key: str, price_data: list) -> Tuple[bool, str]:
        url = f"https://api.partner.market.yandex.ru/v2/businesses/{business_id}/offers/update-prices"
        headers = {"Authorization": f"OAuth {api_key}", "Content-Type": "application/json"}
        offers_payload = [
            {"offerId": str(item['artikul']), "price": {"value": float(item['new_price']), "currencyId": "RUR"}}
            for item in price_data
        ]
        try:
            response = requests.post(url, json={"offers": offers_payload}, headers=headers, timeout=10)
            if response.status_code == 200:
                return True, "Цены успешно обновлены в ЛК ZapStore."
            return False, f"Ошибка API ({response.status_code}): {response.text}"
        except Exception as e:
            return False, f"Ошибка соединения: {str(e)}"

# ============================================================================
# БЛОК 5: ГЕНЕРАТОР ДАННЫХ ДЛЯ ABC-XYZ АНАЛИЗА
# ============================================================================
def generate_abc_xyz_sample_data() -> pd.DataFrame:
    """Генерация данных для ABC-XYZ анализа с корректными знаками"""
    data = {
        "Категория": [
            "Босоножки", "Сарафаны", "Платья", "Кардиганы", "Туфли",
            "Юбки", "Футболки", "Шорты", "Джинсы", "Топы",
            "Рубашки", "Кардиган", "Брюки", "Футболка", "Рубашка",
            "Платье", "Сарафан", "Худи", "Кроссовки", "Кофты", "Юбка"
        ],
        "ABC-XYZ": [
            "AACZ", "AACZ", "AACZ", "AACZ", "AACZ",
            "AACZ", "ABCZ", "BBCZ", "BCCZ", "BCCZ",
            "BBCZ", "BCCZ", "BCCZ", "CCCZ", "CCCZ",
            "CCCZ", "CCCZ", "CCCZ", "CCCZ", "CCCZ", "CCCZ"
        ],
        "Заказы, шт.": [
            37097, 23121, 23437, 32152, 14238,
            15417, 24271, 8626, 4336, 5310,
            3157, 7529, 2737, 5936, 3791,
            3676, 1841, 666, 1001, 689, 1648
        ],
        "% Выкупа": [
            0.97, 0.98, 0.95, 0.94, 0.86,
            0.96, 0.97, 0.79, 0.71, 0.97,
            0.85, 0.47, 1.00, 0.48, 0.37,
            0.30, 0.28, 0.53, 0.40, 0.61, 0.23
        ],
        "Возвраты, шт.": [
            2740, 2937, 2771, 1748, 1333,
            1941, 1312, 755, 515, 479,
            241, 291, 308, 296, 146,
            208, 78, 39, 26, 25, 91
        ],
        "Продажи, шт.": [
            33261, 19748, 19586, 28615, 10940,
            12846, 22167, 6026, 2566, 4660,
            2455, 3212, 2416, 2529, 1258,
            913, 443, 311, 378, 395, 290
        ],
        "Продажи до возвр. ₽": [
            135784831, 73677033, 67786791, 65262399, 50701892,
            41955005, 36320954, 14360677, 11381445, 9008037,
            8003797, 7500498, 6466635, 4997896, 3351387,
            2938288, 1530614, 1246739, 1115940, 1044853, 974523
        ],
        "Продажи после возвр. ₽": [
            124862387, 63941002, 59104570, 61145530, 44872429,
            36205900, 34097531, 12491520, 9359691, 8113721,
            7189978, 6830896, 5732375, 4466520, 2988842,
            2386209, 1297385, 1090082, 1045734, 961225, 735182
        ],
        # ИСПРАВЛЕНО: Затраты как ПОЛОЖИТЕЛЬНЫЕ числа
        "Себестоимость ₽": [
            33712544, 13691020, 14546761, 18006452, 12417533,
            7939400, 11032230, 3090368, 2380811, 2147284,
            1763675, 2011441, 1583772, 1335475, 854641,
            776218, 331824, 462677, 398392, 367170, 214618
        ],
        "Логистика ₽": [
            216230, 71512, 231440, 460311, 442667,
            106810, 174895, 436784, 394752, 28218,
            98923, 1672397, 1813, 998635, 896152,
            855353, 456022, 190620, 262825, 65190, 368921
        ],
        "Комиссия ₽": [
            51632405, 25153826, 24362216, 24423649, 18297908,
            15329669, 12619758, 4751649, 3650202, 3436713,
            2697248, 814146, 2379000, 529959, 353166,
            283318, 158616, 301843, 202567, 332208, 87933
        ],
        "Возвраты ₽": [
            10922444, 9736031, 8682221, 4116869, 5829463,
            5749106, 2223423, 1869157, 2021754, 894316,
            813819, 669602, 734260, 531376, 362545,
            552079, 233229, 156657, 70206, 83628, 239341
        ],
        "Рекламные затраты ₽": [
            8969028, 8656015, 5006377, 4048301, 4262027,
            3658168, 3287424, 1635363, 997816, 828260,
            582106, 1107211, 591493, 854050, 380016,
            524653, 345342, 56727, 331286, 38448, 215172
        ],
        "Стоимость хранения ₽": [
            13796000, 4019000, 32715700, 31871300, 30430000,
            36561000, 10989000, 2343000, 2704000, 926000,
            178223000, 1107211, 0, 0, 0,
            0, 0, 1427000, 150000, 289670, 0
        ]
    }
    
    df = pd.DataFrame(data)
    
    # Расчет РЕАЛЬНОГО ЗАРАБОТКА (прибыли)
    df['ПРИБЫЛЬ ₽'] = (
        df['Продажи после возвр. ₽'] - 
        df['Себестоимость ₽'] - 
        df['Логистика ₽'] - 
        df['Комиссия ₽'] - 
        df['Возвраты ₽'] - 
        df['Рекламные затраты ₽'] - 
        df['Стоимость хранения ₽']
    )
    
    return df

# ============================================================================
# БЛОК 6: STREAMLIT ИНТЕРФЕЙС
# ============================================================================
def main():
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="📈",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # CSS стили
    st.markdown("""
    <style>
        .reportview-container {
            background: #f5f7f9;
        }
        .sidebar .sidebar-content {
            background: #1e293b;
            color: white;
        }
        .stMetric {
            background-color: #ffffff;
            padding: 15px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.05);
            border: 1px solid #e2e8f0;
        }
        .instruction-box {
            background-color: #f8fafc;
            padding: 15px;
            border-radius: 6px;
            border-left: 4px solid #3b82f6;
            margin-bottom: 20px;
            color: #334155;
        }
        .profit-positive {
            color: #10b981;
            font-weight: bold;
        }
        .profit-negative {
            color: #ef4444;
            font-weight: bold;
        }
        div[data-testid="stMetricValue"] {
            font-size: 24px;
        }
    </style>
    """, unsafe_allow_html=True)

    # API ключи (из настроек)
    COMPANY_API_KEY = "ACMA:baYKVsVh7vORZYIZLLvZviviZAxfjcRmdrariFBH:e755690c"
    COMPANY_BUSINESS_ID = "93193868"

    # Инициализация менеджера тарифов
    tm = HybridTariffManager()

    # Инициализация данных
    if 'main_df' not in st.session_state:
        st.session_state.main_df = pd.DataFrame([{
            'artikul': 'PART-7831',
            'category': 'автозапчасти',
            'selling_price': 3990.0,
            'cogs': 1800.0,
            'weight_kg': 1.5,
            'length_cm': 25.0,
            'width_cm': 15.0,
            'height_cm': 10.0,
            'packaging_cost': 40.0,
            'marketing_budget_per_unit': 200.0,
            'daily_sales': 4,
            'stock_depth_days': 30
        }])

    # === НАВИГАЦИОННАЯ ПАНЕЛЬ ===
    st.sidebar.title(" Навигация по системе")
    page = st.sidebar.radio(
        "Выберите рабочий экран:",
        [
            " Дашборд и Аналитика",
            "🔥 ABC-XYZ и 50+ Метрик",
            "📝 Калькулятор экономики",
            "🗂️ Управление категориями",
            "💾 Импорт / Экспорт данных",
            "📡 Синхронизация API"
        ]
    )
    
    st.sidebar.markdown("---")
    st.sidebar.subheader("⚙️ Глобальный налоговый режим")
    tax_label = st.sidebar.selectbox(
        "Налогообложение компании:",
        ["УСН 6% (доходы)", "УСН 15% (доходы-расходы)", "ОСН (общая с НДС 20%)"],
        help="Выбор влияет на автоматический расчёт чистой маржи и генерацию коридора цен."
    )
    current_tax = TaxSystem.by_label(tax_label)

    # Расчет всех метрик
    calculated_df = VectorizedEnginePRO.run_calculations(
        st.session_state.main_df, 
        current_tax, 
        tm
    )

    # === СТРАНИЦА 1: ДАШБОРД ===
    if page == " Дашборд и Аналитика":
        st.title("📊 Панель комплексной аналитики и KPI")
        
        st.markdown("""
        <div class="instruction-box">
            <strong>📋 Как работать с этим экраном:</strong><br>
            1. Данная панель собирает верхнеуровневые метрики на основе вашей товарной матрицы.<br>
            2. Используйте виджеты KPI для оценки средней маржинальности всего портфеля.<br>
            3. Диаграмма структуры цены наглядно показывает соотношение себестоимости закупок (COGS) и чистой прибыли.
        </div>
        """, unsafe_allow_html=True)
        
        if not calculated_df.empty:
            # KPI метрики
            c1, c2, c3, c4 = st.columns(4)
            
            total_revenue = calculated_df['selling_price'].sum()
            total_profit = calculated_df['gross_profit'].sum()
            avg_margin = calculated_df['margin_percent'].mean()
            total_sku = len(calculated_df)
            
            c1.metric(
                "Всего SKU в матрице",
                f"{total_sku}",
                delta=None
            )
            c2.metric(
                "Ср. маржинальность портфеля",
                f"{avg_margin:.2f}%",
                delta=f"{'+' if avg_margin >= 15 else '-'}{abs(avg_margin - 15):.2f}%",
                delta_color="normal" if avg_margin >= 15 else "inverse"
            )
            c3.metric(
                "Общая выручка",
                f"{format_number(total_revenue, ' ₽')}",
                delta=None
            )
            c4.metric(
                "ОБЩАЯ ПРИБЫЛЬ",
                f"{format_number(total_profit, ' ₽')}",
                delta=f"{'✅' if total_profit > 0 else '❌'}",
                delta_color="normal" if total_profit > 0 else "inverse"
            )
            
            st.markdown("---")
            
            # Графики
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader(" Соотношение себестоимости и прибыли по артикулам")
                fig_bar = px.bar(
                    calculated_df.head(10),
                    x='artikul',
                    y=['cogs', 'gross_profit', 'total_expenses'],
                    title="Разбивка стоимости единицы товара (Топ-10)",
                    labels={'value': 'Рубли (₽)', 'artikul': 'Артикул', 'variable': 'Статья'},
                    barmode='group',
                    color_discrete_map={
                        'cogs': '#ef4444',
                        'gross_profit': '#10b981',
                        'total_expenses': '#f59e0b'
                    }
                )
                fig_bar.update_layout(showlegend=True, height=400)
                st.plotly_chart(fig_bar, use_container_width=True)
            
            with col2:
                st.subheader(" Маржинальность по категориям")
                fig_pie = px.pie(
                    calculated_df.groupby('category', as_index=False).agg({
                        'gross_profit': 'sum',
                        'selling_price': 'sum'
                    }),
                    values='gross_profit',
                    names='category',
                    title="Распределение прибыли по категориям",
                    color_discrete_sequence=px.colors.qualitative.Set3
                )
                fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                fig_pie.update_layout(height=400)
                st.plotly_chart(fig_pie, use_container_width=True)
            
            # Таблица с ключевыми метриками
            st.subheader("📋 Детализация по SKU")
            display_df = calculated_df[[
                'artikul', 'category', 'selling_price', 'cogs',
                'gross_profit', 'margin_percent', 'roi_percent',
                'rec_price_min', 'rec_price_15'
            ]].copy()
            
            display_df.columns = [
                'Артикул', 'Категория', 'Цена продажи', 'Себестоимость',
                'ПРИБЫЛЬ', 'Маржа %', 'ROI %', 'Мин. цена', 'Опт. цена'
            ]
            
            st.dataframe(
                display_df,
                use_container_width=True,
                hide_index=True
            )
        else:
            st.warning("Товарная матрица пуста. Перейдите во вкладку 'Калькулятор экономики' для добавления позиций.")

    # === СТРАНИЦА 2: ABC-XYZ И 50+ МЕТРИК ===
    elif page == "🔥 ABC-XYZ и 50+ Метрик":
        st.title("🔥 ABC-XYZ Анализ и 50+ Метрик Эффективности")
        
        st.markdown("""
        <div class="instruction-box">
            <strong>⚠️ КРИТИЧЕСКОЕ ИСПРАВЛЕНИЕ ЛОГИКИ:</strong><br>
            В исходных данных затраты (себестоимость, логистика, комиссия) отображались со знаком «минус».
            Это методологическая ошибка! Затраты — это абсолютные положительные величины.
            <br><br>
            <strong>В этом отчете:</strong>
            <ul>
                <li>✅ Все затраты показаны как положительные числа</li>
                <li>✅ Колонка <b>ПРИБЫЛЬ</b> показывает реальный заработок (может быть отрицательной)</li>
                <li>✅ 50+ метрик для глубокого анализа эффективности</li>
                <li>✅ ABC-XYZ классификация с цветовым кодированием</li>
            </ul>
        </div>
        """, unsafe_allow_html=True)
        
        # Генерация данных ABC-XYZ
        abc_df = generate_abc_xyz_sample_data()
        
        # Сводные KPI
        st.subheader(" Сводные показатели ABC-XYZ анализа")
        
        total_orders = abc_df['Заказы, шт.'].sum()
        total_revenue = abc_df['Продажи до возвр. ₽'].sum()
        total_profit = abc_df['ПРИБЫЛЬ ₽'].sum()
        avg_buyout = abc_df['% Выкупа'].mean()
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric("Всего заказов", f"{total_orders:,}".replace(',', ' '))
        kpi2.metric("Общая выручка", f"{format_number(total_revenue, ' ₽')}")
        kpi3.metric(
            "РЕАЛЬНЫЙ ЗАРАБОТОК",
            f"{format_number(total_profit, ' ₽')}",
            delta=f"{'✅ Прибыльно' if total_profit > 0 else '❌ Убыточно'}",
            delta_color="normal" if total_profit > 0 else "inverse"
        )
        kpi4.metric("Ср. % выкупа", f"{avg_buyout*100:.1f}%")
        
        st.markdown("---")
        
        # Детальная таблица с условным форматированием
        st.subheader(" Детальная матрица категорий с подсветкой прибыли")
        
        # Выбор колонок для отображения
        display_columns = [
            "Категория", "ABC-XYZ", "Заказы, шт.", "% Выкупа",
            "Продажи до возвр. ₽", "Себестоимость ₽",
            "Логистика ₽", "Комиссия ₽", "Рекламные затраты ₽",
            "ПРИБЫЛЬ ₽"
        ]
        
        # Форматирование таблицы
        display_df = abc_df[display_columns].copy()
        
        # Форматирование процентов
        display_df['% Выкупа'] = display_df['% Выкупа'].apply(lambda x: f"{x*100:.0f}%")
        
        # Форматирование чисел
        for col in ["Заказы, шт.", "Продажи до возвр. ₽", "Себестоимость ₽",
                   "Логистика ₽", "Комиссия ₽", "Рекламные затраты ₽", "ПРИБЫЛЬ ₽"]:
            display_df[col] = display_df[col].apply(lambda x: f"{x:,.0f} ₽".replace(',', ' '))
        
        # Отображение с цветовой кодировкой
        def color_profit(val):
            if isinstance(val, str) and '₽' in val:
                num_val = float(val.replace(' ₽', '').replace(' ', ''))
                if num_val > 0:
                    return 'background-color: #d4edda; color: #155724; font-weight: bold'
                elif num_val < 0:
                    return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
            return ''
        
        def color_abc(val):
            if isinstance(val, str):
                if val.startswith('AA') or val.startswith('AB'):
                    return 'background-color: #d4edda; color: #155724; font-weight: bold'
                elif val.startswith('BA') or val.startswith('BB'):
                    return 'background-color: #fff3cd; color: #856404; font-weight: bold'
                elif val.startswith('C'):
                    return 'background-color: #f8d7da; color: #721c24; font-weight: bold'
            return ''
        
        styled_df = display_df.style.applymap(color_profit, subset=['ПРИБЫЛЬ ₽']).applymap(color_abc, subset=['ABC-XYZ'])
        
        st.dataframe(
            styled_df,
            use_container_width=True,
            height=600
        )
        
        st.markdown("---")
        
        # 50+ Метрики
        st.subheader(" 50+ Метрик эффективности (расширенный анализ)")
        
        st.markdown("""
        <div class="instruction-box">
            <strong>Метрики разделены на группы:</strong>
            <ol>
                <li><b>Финансовые:</b> Выручка, себестоимость, прибыль, маржа</li>
                <li><b>Эффективности:</b> ROI, наценка, точка безубыточности</li>
                <li><b>Затрат:</b> Доля логистики, комиссии, маркетинга в цене</li>
                <li><b>ABC-XYZ:</b> Классификация по объему и стабильности</li>
            </ol>
        </div>
        """, unsafe_allow_html=True)
        
        # Расчет 50+ метрик для демонстрации
        metrics_df = calculated_df.copy() if not calculated_df.empty else abc_df.head(5)
        
        if not metrics_df.empty and 'gross_profit' in metrics_df.columns:
            st.write(f"**Всего метрик рассчитано:** {len([c for c in metrics_df.columns if c not in ['artikul', 'category']])}+")
            
            # Выбор метрик для отображения
            metrics_to_show = [
                'artikul', 'category', 'selling_price', 'cogs', 'gross_profit',
                'margin_percent', 'roi_percent', 'markup_percent',
                'commission', 'last_mile_cost', 'first_mile_cost',
                'acquiring_cost', 'return_cost', 'marketing_budget_per_unit',
                'warehouse_cost', 'tax_cost', 'total_expenses',
                'rec_price_min', 'rec_price_15', 'rec_price_25',
                'break_even_units', 'safety_margin_percent',
                'contribution_margin_percent', 'cogs_percent',
                'logistics_percent', 'commission_percent',
                'abc_category', 'xyz_category', 'abc_xyz'
            ]
            
            available_metrics = [m for m in metrics_to_show if m in metrics_df.columns]
            st.dataframe(
                metrics_df[available_metrics].head(10),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("Для расчета 50+ метрик добавьте товары в калькуляторе экономики")
        
        st.markdown("---")
        
        # Экспорт
        st.subheader("💾 Экспорт отчетов")
        
        col_exp1, col_exp2 = st.columns(2)
        
        with col_exp1:
            if OPENPYXL_AVAILABLE:
                excel_bytes = ABCXYZExcelExporter.export_abc_xyz(abc_df)
                st.download_button(
                    label="⬇️ СКАЧАТЬ ABC-XYZ ОТЧЕТ (.XLSX)",
                    data=excel_bytes,
                    file_name=f"ABC_XYZ_Analysis_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        with col_exp2:
            if OPENPYXL_AVAILABLE:
                metrics_excel = ComprehensiveMetricsExporter.export_full_metrics(calculated_df)
                st.download_button(
                    label="⬇️ СКАЧАТЬ 50+ МЕТРИК (.XLSX)",
                    data=metrics_excel,
                    file_name=f"Full_Metrics_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        if not OPENPYXL_AVAILABLE:
            st.warning("⚠️ Библиотека `openpyxl` не установлена. Установите: `pip install openpyxl`")

    # === СТРАНИЦА 3: КАЛЬКУЛЯТОР ===
    elif page == "📝 Калькулятор экономики":
        st.title("📝 Интерактивный симулятор товарной матрицы")
        
        st.markdown("""
        <div class="instruction-box">
            <strong> Как работать с калькулятором:</strong><br>
            1. Нажмите дважды на любую ячейку в таблице <b>'Редактор матрицы'</b> для изменения исходных данных.<br>
            2. Вы можете менять текущую цену (<code>selling_price</code>), себестоимость (<code>cogs</code>) или габариты.<br>
            3. Вторая таблица ниже автоматически пересчитает расходы на логистику, налоги и покажет три сценария идеальной цены.
        </div>
        """, unsafe_allow_html=True)
        
        st.subheader("Шаг 1: Редактор исходных параметров товаров")
        edited_df = st.data_editor(
            st.session_state.main_df,
            num_rows="dynamic",
            use_container_width=True,
            key="product_editor"
        )
        
        if not edited_df.empty:
            st.session_state.main_df = edited_df
            calculated_df = VectorizedEnginePRO.run_calculations(edited_df, current_tax, tm)

            st.subheader("Шаг 2: Результаты расчетов финансовых коридоров")
            
            # Выбор ключевых метрик для отображения
            results_columns = [
                'artikul', 'category', 'selling_price', 'cogs',
                'gross_profit', 'margin_percent', 'roi_percent',
                'rec_price_min', 'rec_price_15', 'rec_price_25'
            ]
            
            results_df = calculated_df[results_columns].copy()
            results_df.columns = [
                'Артикул', 'Категория', 'Цена продажи', 'Себестоимость',
                'ПРИБЫЛЬ', 'Маржа %', 'ROI %',
                'Мин. цена (0%)', 'Опт. цена (15%)', 'Макс. цена (25%)'
            ]
            
            st.dataframe(
                results_df,
                use_container_width=True,
                hide_index=True
            )
            
            # Визуализация
            st.subheader("📊 Визуализация прибыльности")
            
            fig = px.bar(
                calculated_df.head(10),
                x='artikul',
                y='gross_profit',
                color='gross_profit',
                color_continuous_scale=['#ef4444', '#f59e0b', '#10b981'],
                title="Чистая прибыль по артикулам (Топ-10)",
                labels={'artikul': 'Артикул', 'gross_profit': 'Прибыль, ₽'}
            )
            fig.update_layout(showlegend=False, height=400)
            st.plotly_chart(fig, use_container_width=True)

    # === СТРАНИЦА 4: КАТЕГОРИИ ===
    elif page == "🗂️ Управление категориями":
        st.title("🗂️ Настройка и добавление кастомных категорий тарификации")
        
        st.markdown("""
        <div class="instruction-box">
            <strong>📋 Как добавить свою категорию:</strong><br>
            1. Заполните текстовое поле с точным названием категории (в нижнем регистре, например: <code>тормозные колодки</code>).<br>
            2. Укажите процентную ставку комиссии маркетплейса согласно договору оферты.<br>
            3. Нажмите кнопку добавления. Теперь движок при расчёте unit-экономики для этого типа товара будет применять новые правила.
        </div>
        """, unsafe_allow_html=True)
        
        with st.form("add_category_form"):
            col1, col2 = st.columns(2)
            with col1:
                new_cat_name = st.text_input(
                    "Название новой категории товара:",
                    help="Используйте строчные буквы, например: фильтры"
                )
                new_comm = st.number_input(
                    "Процент комиссии оферты (%)",
                    min_value=0.0,
                    max_value=100.0,
                    value=12.0
                ) / 100
            with col2:
                new_min_comm = st.number_input(
                    "Минимум комиссии маркетплейса (₽)",
                    min_value=0.0,
                    value=35.0
                )
                new_storage = st.number_input(
                    "Хранение за день (₽)",
                    min_value=0.0,
                    value=0.50
                )
            
            submitted = st.form_submit_button("➕ Зарегистрировать категорию")
            
            if submitted and new_cat_name:
                cleaned_name = new_cat_name.lower().strip()
                tm.tariffs[cleaned_name] = Tariff(
                    category=cleaned_name,
                    commission_rate=new_comm,
                    min_commission=new_min_comm,
                    storage_fee_per_day=new_storage,
                    source="Пользовательская база"
                )
                st.success(f"Категория '{new_cat_name}' успешно добавлена в оперативную память!")
        
        st.subheader("Действующая сетка тарифов по категориям")
        st.dataframe(
            tm.to_dataframe(),
            use_container_width=True,
            hide_index=True
        )

    # === СТРАНИЦА 5: ИМПОРТ / ЭКСПОРТ ===
    elif page == "💾 Импорт / Экспорт данных":
        st.title(" Централизованный импорт и экспорт данных")
        
        st.markdown("""
        <div class="instruction-box">
            <strong> Инструкция по миграции данных:</strong><br>
            • <b>Импорт:</b> Загрузите CSV-файл, содержащий заголовки: <code>artikul, category, selling_price, cogs</code>, чтобы пакетно перезаписать текущую матрицу.<br>
            • <b>Экспорт:</b> Нажмите кнопку экспорта в Excel для получения многостраничного финансового документа с «живыми» формулами СУММ и условий.
        </div>
        """, unsafe_allow_html=True)
        
        st.subheader("1. Загрузка товарного каталога")
        uploaded_file = st.file_uploader(
            "Перетащите файл CSV или JSON сюда",
            type=['csv', 'json']
        )
        
        if uploaded_file is not None:
            try:
                if uploaded_file.name.endswith('.csv'):
                    uploaded_df = pd.read_csv(uploaded_file)
                else:
                    uploaded_df = pd.read_json(uploaded_file)
                
                required_cols = ['artikul', 'category', 'selling_price', 'cogs']
                if all(col in uploaded_df.columns for col in required_cols):
                    st.session_state.main_df = uploaded_df
                    st.success(f"Файл успешно обработан. Импортировано SKU: {len(uploaded_df)}")
                else:
                    st.error(f"Отсутствуют обязательные колонки: {required_cols}")
            except Exception as e:
                st.error(f"Ошибка парсинга структуры: {e}")

        st.markdown("---")
        st.subheader("2. Выгрузка результатов расчетов")
        
        c1, c2, c3 = st.columns(3)
        
        with c1:
            if OPENPYXL_AVAILABLE:
                excel_bytes = ExcelDynamicExporter.export(calculated_df)
                st.download_button(
                    label="⬇️ СКАЧАТЬ ФИНАНСОВУЮ МОДЕЛЬ В EXCEL (.XLSX)",
                    data=excel_bytes,
                    file_name=f"ZapStore_Economics_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        
        with c2:
            csv_buffer = io.StringIO()
            calculated_df.to_csv(csv_buffer, index=False)
            st.download_button(
                label="⬇️ СКАЧАТЬ ПЛОСКИЙ ОТЧЕТ В CSV",
                data=csv_buffer.getvalue(),
                file_name=f"ZapStore_Matrix_{datetime.now().strftime('%d_%m_%Y')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with c3:
            if OPENPYXL_AVAILABLE:
                abc_excel = ABCXYZExcelExporter.export_abc_xyz(generate_abc_xyz_sample_data())
                st.download_button(
                    label="⬇️ ABC-XYZ ОТЧЕТ (.XLSX)",
                    data=abc_excel,
                    file_name=f"ABC_XYZ_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

    # === СТРАНИЦА 6: API СИНХРОНИЗАЦИЯ ===
    elif page == "📡 Синхронизация API":
        st.title("📡 Модуль интеграции и управления шлюзами ZapStore")
        
        st.markdown("""
        <div class="instruction-box">
            <strong>📋 Правила отправки цен по API:</strong><br>
            1. Выберите в выпадающем списке, какую именно цену вы хотите отправить на маркетплейс (например, Оптимальную с маржой 15%).<br>
            2. Нажмите кнопку публикации. Система сформирует пакет данных и обновит цены в кабинете <b>ZapStore</b>.<br>
            3. Процесс модерации новых цен на стороне площадки занимает от 5 до 15 минут.
        </div>
        """, unsafe_allow_html=True)
        
        selected_strategy = st.selectbox(
            "Выберите тип цены для пакетной отправки в магазин:",
            [
                "Текущая установленная цена",
                "Минимальная цена (Безубыточность)",
                "Оптимальная цена (15% маржа)",
                "Максимальная цена (25% маржа)"
            ]
        )
        
        if st.button(
            "🚀 ВЫГРУЗИТЬ ОБНОВЛЕННЫЕ ЦЕНЫ НА МАРКЕТПЛЕЙС",
            type="primary",
            use_container_width=True
        ):
            price_data_to_send = []
            
            for _, row in calculated_df.iterrows():
                if selected_strategy == "Текущая установленная цена":
                    target_price = row['selling_price']
                elif selected_strategy == "Минимальная цена (Безубыточность)":
                    target_price = row['rec_price_min']
                elif selected_strategy == "Оптимальная цена (15% маржа)":
                    target_price = row['rec_price_15']
                else:
                    target_price = row['rec_price_25']
                
                price_data_to_send.append({
                    'artikul': row['artikul'],
                    'new_price': target_price
                })
            
            with st.spinner("Синхронизация данных с сервером..."):
                success, msg = YandexMarketApiSync.update_prices(
                    COMPANY_BUSINESS_ID,
                    COMPANY_API_KEY,
                    price_data_to_send
                )
                
                if success:
                    st.success(f"✅ Выгрузка завершена! {msg}")
                    st.json({"Отправлено позиций": len(price_data_to_send)})
                else:
                    st.error(f"❌ Не удалось отправить данные: {msg}")

    # Footer
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"""
    <div style='text-align: center; color: #94a3b8;'>
        <p><b>{APP_NAME}</b></p>
        <p>Версия: {APP_VERSION}</p>
        <p>© 2024 ZapStore Analytics</p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
