"""
================================================================================
🚗 ЮНИТ-ЭКОНОМИКА ДЛЯ АВТОЗАПЧАСТЕЙ (FBS ЯНДЕКС МАРКЕТ) - ПРОДУКТИВНАЯ ВЕРСИЯ
================================================================================
✅ Только Яндекс Маркет FBS (единый тариф)
✅ Именованные диапазоны в Excel (автообновление)
✅ Настраиваемые спецрасходы через интерфейс
✅ Кэширование загруженных данных
✅ Интерактивные графики (plotly)
✅ Ключевые метрики на главной
✅ Обработка ошибок API и валидация
✅ Условное форматирование в Excel
================================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import requests
import json
import os
import io
import math
import traceback
import logging
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, asdict
from decimal import Decimal, ROUND_HALF_UP

# Графика
import plotly.express as px
import plotly.graph_objects as go

# Excel
from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# -----------------------------------------------------------------------------
# НАСТРОЙКА ЛОГГИРОВАНИЯ
# -----------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('UnitEconomicsFBS')

# -----------------------------------------------------------------------------
# КОНСТАНТЫ
# -----------------------------------------------------------------------------
APP_VERSION = "3.0.0"
APP_NAME = "🚗 Юнит-экономика (FBS Яндекс Маркет)"
BASE_DIR = Path(__file__).parent.resolve() if '__file__' in locals() else Path.cwd()
DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
TEMP_DIR = BASE_DIR / "temp"
for d in [DATA_DIR, CACHE_DIR, TEMP_DIR]:
    d.mkdir(parents=True, exist_ok=True)

TARIFF_JSON = DATA_DIR / "yandex_tariff.json"
EXCEL_ROW_LIMIT_FOR_FORMULAS = 10000  # при большем количестве используем статику

# -----------------------------------------------------------------------------
# УТИЛИТЫ
# -----------------------------------------------------------------------------
def money_round(value: float, decimals: int = 2) -> float:
    return float(Decimal(str(value)).quantize(Decimal(f"0.{'0' * decimals}"), rounding=ROUND_HALF_UP))

def safe_float(val) -> float:
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0

# -----------------------------------------------------------------------------
# МОДЕЛЬ ТАРИФА (только Яндекс Маркет)
# -----------------------------------------------------------------------------
@dataclass
class YandexTariff:
    commission_rate: float = 0.14          # доля
    min_commission: float = 0.0            # руб
    logistics_base: float = 45.0           # руб за отправление
    logistics_per_kg: float = 14.0         # руб за кг
    storage_per_day_per_liter: float = 0.25 # руб за литр в день
    acquiring_fee: float = 0.02            # доля
    return_fee: float = 0.02               # доля (резерв)

    def to_dict(self) -> Dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict) -> 'YandexTariff':
        return cls(**data)

    def save(self):
        with open(TARIFF_JSON, 'w', encoding='utf-8') as f:
            json.dump(self.to_dict(), f, ensure_ascii=False, indent=2)

    @classmethod
    def load(cls) -> 'YandexTariff':
        if TARIFF_JSON.exists():
            try:
                with open(TARIFF_JSON, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                return cls.from_dict(data)
            except Exception as e:
                logger.warning(f"Ошибка загрузки тарифа: {e}, используем дефолт")
        return cls()  # дефолтный

# -----------------------------------------------------------------------------
# МЕНЕДЖЕР ТАРИФА (загрузка из CSV и API)
# -----------------------------------------------------------------------------
class TariffManager:
    def __init__(self):
        self.tariff = YandexTariff.load()

    def update_from_csv(self, file_content: bytes) -> Tuple[bool, str]:
        """Загружает тариф из CSV с колонками: commission_rate, min_commission, ..."""
        try:
            df = pd.read_csv(io.BytesIO(file_content), sep=';')
            required = ['commission_rate', 'logistics_base', 'logistics_per_kg', 'storage_per_day']
            if not all(c in df.columns for c in required):
                return False, f"Отсутствуют колонки: {required}"
            row = df.iloc[0]
            self.tariff.commission_rate = float(row.get('commission_rate', 0.14))
            self.tariff.min_commission = float(row.get('min_commission', 0.0))
            self.tariff.logistics_base = float(row.get('logistics_base', 45.0))
            self.tariff.logistics_per_kg = float(row.get('logistics_per_kg', 14.0))
            self.tariff.storage_per_day_per_liter = float(row.get('storage_per_day', 0.25))
            self.tariff.acquiring_fee = float(row.get('acquiring_fee', 0.02))
            self.tariff.return_fee = float(row.get('return_fee', 0.02))
            self.tariff.save()
            return True, f"Тариф обновлён из CSV"
        except Exception as e:
            return False, f"Ошибка парсинга CSV: {e}"

    def fetch_from_yandex_api(self, oauth_token: str, campaign_id: int) -> Tuple[bool, str]:
        """Загружает тариф через API Яндекс Маркета (усреднение по категориям)"""
        url = f"https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/categories/commissions"
        headers = {"Authorization": f"OAuth {oauth_token}", "Accept": "application/json"}
        try:
            resp = requests.get(url, headers=headers, timeout=15)
            if resp.status_code != 200:
                return False, f"API вернул {resp.status_code}: {resp.text[:200]}"
            data = resp.json()
            categories = data.get("result", {}).get("categories", [])
            if not categories:
                return False, "API вернул пустой список категорий. Проверьте права доступа."
            
            total_comm = 0.0
            total_log = 0.0
            count = 0
            for cat in categories:
                total_comm += cat.get("commissionPercent", 14) / 100.0
                total_log += cat.get("fbsLogisticsBase", 45.0)
                count += 1
            avg_comm = total_comm / count
            avg_log = total_log / count

            self.tariff.commission_rate = avg_comm
            self.tariff.logistics_base = avg_log
            # остальные параметры оставляем прежними (или можно попробовать вытащить из API, но их нет)
            self.tariff.save()
            return True, f"Тариф обновлён из API (усреднено по {count} категориям)"
        except Exception as e:
            return False, f"Ошибка запроса: {e}"

    def get_tariff_dict(self) -> Dict:
        return self.tariff.to_dict()

    def get_current_tariff_csv(self) -> bytes:
        """Возвращает CSV с текущим тарифом для скачивания"""
        df = pd.DataFrame([self.tariff.to_dict()])
        return df.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

# -----------------------------------------------------------------------------
# СПРАВОЧНИК КАТЕГОРИЙ (дефолтные габариты и флаги)
# -----------------------------------------------------------------------------
class AutoPartsCategoriesDB:
    def __init__(self):
        self.db = {
            "фильтры": (1.5, 0.5, False, False),
            "масла": (5.0, 4.0, True, False),
            "колодки": (0.8, 1.2, False, False),
            "диски": (3.0, 4.0, False, True),
            "амортизаторы": (4.0, 3.5, False, True),
            "аккумуляторы": (12.0, 15.0, True, True),
            "шины": (25.0, 10.0, False, False),
            "фары": (6.0, 2.5, False, True),
            "двигатель": (50.0, 80.0, True, True),
            "кпп": (40.0, 50.0, True, True),
        }

    def get_defaults(self, category: str) -> Dict:
        cat_key = category.lower().strip()
        for key, vals in self.db.items():
            if key in cat_key:
                return {"volume_l": vals[0], "weight_kg": vals[1],
                        "is_hazardous": vals[2], "is_fragile": vals[3]}
        return {"volume_l": 2.0, "weight_kg": 1.0,
                "is_hazardous": False, "is_fragile": False}

# -----------------------------------------------------------------------------
# ВАЛИДАТОР ГАБАРИТОВ И РАСЧЁТ ОПЛАЧИВАЕМОГО ВЕСА
# -----------------------------------------------------------------------------
class FBSDimensionsValidator:
    @staticmethod
    def normalize_dimension(value: float, unit_hint: str = "cm") -> float:
        if not value or value <= 0:
            return 0.0
        unit = str(unit_hint).lower()
        if any(x in unit for x in ['mm', 'мм']):
            return value / 10.0
        if any(x in unit for x in ['m', 'метр']):
            return value * 100.0
        if value > 300:  # вероятно, в мм
            return value / 10.0
        return value

    @staticmethod
    def calculate_billable_weight(weight_kg: float, length_cm: float, width_cm: float, height_cm: float) -> float:
        if length_cm <= 0 or width_cm <= 0 or height_cm <= 0:
            return max(0.1, weight_kg)
        volumetric_weight = (length_cm * width_cm * height_cm) / 5000.0
        billable = max(weight_kg, volumetric_weight)
        return math.ceil(billable * 2) / 2  # округление вверх до 0.5 кг

    @staticmethod
    def validate_batch(df: pd.DataFrame, categories_db: AutoPartsCategoriesDB) -> pd.DataFrame:
        df = df.copy()
        # Нормализация размеров
        for col in ['Длина', 'Ширина', 'Высота']:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: FBSDimensionsValidator.normalize_dimension(safe_float(x)))
        if 'Вес_кг' in df.columns:
            df['Вес_кг'] = df['Вес_кг'].apply(lambda x: safe_float(x) if safe_float(x) < 100 else safe_float(x)/1000)
        # Заполнение пропусков из категорий
        if 'Категория' in df.columns:
            for idx, row in df.iterrows():
                cat = str(row.get('Категория', ''))
                defaults = categories_db.get_defaults(cat)
                if safe_float(df.at[idx, 'Объем_л']) == 0 or pd.isna(df.at[idx, 'Объем_л']):
                    df.at[idx, 'Объем_л'] = defaults['volume_l']
                if safe_float(df.at[idx, 'Вес_кг']) == 0 or pd.isna(df.at[idx, 'Вес_кг']):
                    df.at[idx, 'Вес_кг'] = defaults['weight_kg']
                df.at[idx, 'is_hazardous'] = defaults['is_hazardous']
                df.at[idx, 'is_fragile'] = defaults['is_fragile']
        # Рассчёт оплачиваемого веса
        if all(col in df.columns for col in ['Длина', 'Ширина', 'Высота']):
            df['Оплач_вес'] = df.apply(
                lambda r: FBSDimensionsValidator.calculate_billable_weight(
                    safe_float(r['Вес_кг']),
                    safe_float(r['Длина']),
                    safe_float(r['Ширина']),
                    safe_float(r['Высота'])
                ), axis=1
            )
        else:
            df['Оплач_вес'] = df['Вес_кг']
        return df

# -----------------------------------------------------------------------------
# РАСЧЁТ СПЕЦИФИЧЕСКИХ РАСХОДОВ (с возможностью настройки)
# -----------------------------------------------------------------------------
class FBSSpecificCosts:
    def __init__(self, packaging: float = 45.0, chestny_znak: float = 1.5,
                 labeling: float = 3.0, warranty_reserve: float = 0.02,
                 hazard_surcharge: float = 0.01, fragile_surcharge: float = 0.005):
        self.packaging = packaging
        self.chestny_znak = chestny_znak
        self.labeling = labeling
        self.warranty_reserve = warranty_reserve
        self.hazard_surcharge = hazard_surcharge
        self.fragile_surcharge = fragile_surcharge

    def calculate_batch(self, df: pd.DataFrame) -> pd.DataFrame:
        df = df.copy()
        df['Упаковка_FBS'] = self.packaging
        df['Маркировка'] = self.chestny_znak + self.labeling  # всегда
        df['Гарант_резерв'] = df['Цена'] * self.warranty_reserve
        df['Надбавка_опасный'] = np.where(df.get('is_hazardous', False),
                                          df['Цена'] * self.hazard_surcharge, 0)
        df['Надбавка_хрупкий'] = np.where(df.get('is_fragile', False),
                                          df['Цена'] * self.fragile_surcharge, 0)
        df['Спец_расходы_FBS'] = (df['Упаковка_FBS'] + df['Маркировка'] +
                                  df['Гарант_резерв'] + df['Надбавка_опасный'] +
                                  df['Надбавка_хрупкий'])
        return df

# -----------------------------------------------------------------------------
# ЭКСПОРТ В EXCEL С ИМЕНОВАННЫМИ ДИАПАЗОНАМИ И ДАШБОРДОМ
# -----------------------------------------------------------------------------
class AdvancedExcelExporter:
    def __init__(self, tariff: YandexTariff, specific_costs: FBSSpecificCosts):
        self.tariff = tariff
        self.specific_costs = specific_costs
        self.wb = Workbook()
        self._setup_styles()

    def _setup_styles(self):
        self.header_font = Font(bold=True, color="FFFFFF", size=11)
        self.header_fill = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
        self.input_fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
        self.formula_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        self.border = Border(left=Side(style='thin'), right=Side(style='thin'),
                             top=Side(style='thin'), bottom=Side(style='thin'))

    def _style_header(self, ws, row, max_col):
        for col in range(1, max_col+1):
            cell = ws.cell(row=row, column=col)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = self.border

    def export(self, df: pd.DataFrame, filepath: str):
        # --- Лист с тарифом (одна строка) ---
        ws_tariff = self.wb.create_sheet("Тариф")
        headers = ["commission_rate", "min_commission", "logistics_base", "logistics_per_kg",
                   "storage_per_day", "acquiring_fee", "return_fee"]
        for col, h in enumerate(headers, 1):
            ws_tariff.cell(row=1, column=col, value=h)
        self._style_header(ws_tariff, 1, len(headers))
        row = 2
        ws_tariff.cell(row=row, column=1, value=self.tariff.commission_rate).number_format = '0.00%'
        ws_tariff.cell(row=row, column=2, value=self.tariff.min_commission).number_format = '#,##0.00'
        ws_tariff.cell(row=row, column=3, value=self.tariff.logistics_base).number_format = '#,##0.00'
        ws_tariff.cell(row=row, column=4, value=self.tariff.logistics_per_kg).number_format = '#,##0.00'
        ws_tariff.cell(row=row, column=5, value=self.tariff.storage_per_day_per_liter).number_format = '#,##0.000'
        ws_tariff.cell(row=row, column=6, value=self.tariff.acquiring_fee).number_format = '0.00%'
        ws_tariff.cell(row=row, column=7, value=self.tariff.return_fee).number_format = '0.00%'
        for col in range(1, 8):
            ws_tariff.cell(row=row, column=col).border = self.border
        # Именованный диапазон "TariffRow"
        self.wb.create_named_range("TariffRow", ws_tariff, "$A$2:$G$2")
        for col in range(1, 8):
            ws_tariff.column_dimensions[get_column_letter(col)].width = 18

        # --- Входные данные ---
        ws_in = self.wb.create_sheet("Входные_Данные")
        in_headers = ["Артикул", "Категория", "Цена", "Себестоимость",
                      "Вес_кг", "Длина_см", "Ширина_см", "Высота_см",
                      "Объем_л", "Оплач_вес", "Спец_расходы_FBS"]
        for col, h in enumerate(in_headers, 1):
            ws_in.cell(row=1, column=col, value=h)
        self._style_header(ws_in, 1, len(in_headers))
        # Заполнение данными
        for r_idx, row in df.iterrows():
            excel_row = r_idx + 2
            ws_in.cell(row=excel_row, column=1, value=row.get('Артикул', ''))
            ws_in.cell(row=excel_row, column=2, value=row.get('Категория', ''))
            ws_in.cell(row=excel_row, column=3, value=row.get('Цена', 0)).number_format = '#,##0.00'
            ws_in.cell(row=excel_row, column=4, value=row.get('Себестоимость', 0)).number_format = '#,##0.00'
            ws_in.cell(row=excel_row, column=5, value=row.get('Вес_кг', 0)).number_format = '#,##0.000'
            ws_in.cell(row=excel_row, column=6, value=row.get('Длина', 0)).number_format = '#,##0.0'
            ws_in.cell(row=excel_row, column=7, value=row.get('Ширина', 0)).number_format = '#,##0.0'
            ws_in.cell(row=excel_row, column=8, value=row.get('Высота', 0)).number_format = '#,##0.0'
            ws_in.cell(row=excel_row, column=9, value=row.get('Объем_л', 0)).number_format = '#,##0.000'
            ws_in.cell(row=excel_row, column=10, value=row.get('Оплач_вес', 0)).number_format = '#,##0.000'
            ws_in.cell(row=excel_row, column=11, value=row.get('Спец_расходы_FBS', 0)).number_format = '#,##0.00'
            for col in range(1, 12):
                ws_in.cell(row=excel_row, column=col).fill = self.input_fill
                ws_in.cell(row=excel_row, column=col).border = self.border
        ws_in.freeze_panes = "A2"

        # --- Расчётный лист (с формулами, ссылающимися на TariffRow) ---
        ws_calc = self.wb.create_sheet("Расчет_FBS")
        calc_headers = ["Артикул", "Категория", "Цена", "Себестоимость",
                        "Комиссия_руб", "Логистика_руб", "Хранение_руб",
                        "Эквайринг_руб", "Спец_расходы_FBS",
                        "Итого_расходы", "Прибыль", "Маржа_%"]
        for col, h in enumerate(calc_headers, 1):
            ws_calc.cell(row=1, column=col, value=h)
        self._style_header(ws_calc, 1, len(calc_headers))

        use_formulas = len(df) <= EXCEL_ROW_LIMIT_FOR_FORMULAS
        if not use_formulas:
            st.warning(f"Строк {len(df)} > {EXCEL_ROW_LIMIT_FOR_FORMULAS}. В Excel будут использованы статические значения (без автоматического пересчёта).")

        for r_idx in range(len(df)):
            excel_row = r_idx + 2
            in_row = excel_row   # т.к. строки совпадают

            # Артикул и категория
            ws_calc.cell(row=excel_row, column=1, value=f"=Входные_Данные!A{in_row}")
            ws_calc.cell(row=excel_row, column=2, value=f"=Входные_Данные!B{in_row}")
            # Цена и себестоимость
            ws_calc.cell(row=excel_row, column=3, value=f"=Входные_Данные!C{in_row}").number_format = '#,##0.00'
            ws_calc.cell(row=excel_row, column=4, value=f"=Входные_Данные!D{in_row}").number_format = '#,##0.00'

            if use_formulas:
                # Используем именованный диапазон TariffRow
                comm_formula = f'=MAX(C{excel_row}*INDEX(TariffRow,1,1), INDEX(TariffRow,1,2))'
                log_formula = f'=INDEX(TariffRow,1,3) + Входные_Данные!J{in_row}*INDEX(TariffRow,1,4)'
                stor_formula = f'=Входные_Данные!I{in_row}*INDEX(TariffRow,1,5)*30'
                acq_formula = f'=C{excel_row}*INDEX(TariffRow,1,6)'
                spec_formula = f'=Входные_Данные!K{in_row}'
                total_formula = f'=D{excel_row}+E{excel_row}+F{excel_row}+G{excel_row}+H{excel_row}+I{excel_row}'
                profit_formula = f'=C{excel_row}-J{excel_row}'
                margin_formula = f'=IF(C{excel_row}>0, K{excel_row}/C{excel_row}, 0)'

                ws_calc.cell(row=excel_row, column=5, value=comm_formula).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=6, value=log_formula).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=7, value=stor_formula).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=8, value=acq_formula).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=9, value=spec_formula).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=10, value=total_formula).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=11, value=profit_formula).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=12, value=margin_formula).number_format = '0.00%'
            else:
                # Статический расчёт (для больших данных)
                price = safe_float(df.iloc[r_idx]['Цена'])
                cost = safe_float(df.iloc[r_idx]['Себестоимость'])
                bill_w = safe_float(df.iloc[r_idx]['Оплач_вес'])
                vol = safe_float(df.iloc[r_idx]['Объем_л'])
                spec = safe_float(df.iloc[r_idx]['Спец_расходы_FBS'])
                t = self.tariff
                comm = max(price * t.commission_rate, t.min_commission)
                log = t.logistics_base + bill_w * t.logistics_per_kg
                stor = vol * t.storage_per_day_per_liter * 30
                acq = price * t.acquiring_fee
                total = cost + comm + log + stor + acq + spec
                profit = price - total
                margin = profit / price if price > 0 else 0
                ws_calc.cell(row=excel_row, column=5, value=comm).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=6, value=log).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=7, value=stor).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=8, value=acq).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=9, value=spec).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=10, value=total).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=11, value=profit).number_format = '#,##0.00'
                ws_calc.cell(row=excel_row, column=12, value=margin).number_format = '0.00%'

            for col in range(5, 13):
                cell = ws_calc.cell(row=excel_row, column=col)
                cell.fill = self.formula_fill
                cell.border = self.border

        # Условное форматирование
        last_row = len(df) + 1
        # Маржа < 0 -> красный фон
        ws_calc.conditional_formatting.add(f'L2:L{last_row}',
            CellIsRule(operator='lessThan', formula=['0'],
                       fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')))
        # Маржа >= 15% -> зелёный
        ws_calc.conditional_formatting.add(f'L2:L{last_row}',
            CellIsRule(operator='greaterThanOrEqual', formula=['0.15'],
                       fill=PatternFill(start_color='C6EFCE', end_color='C6EFCE', fill_type='solid')))
        # Прибыль < 0 -> красный
        ws_calc.conditional_formatting.add(f'K2:K{last_row}',
            CellIsRule(operator='lessThan', formula=['0'],
                       fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')))

        for col in range(1, 13):
            ws_calc.column_dimensions[get_column_letter(col)].width = 16
        ws_calc.freeze_panes = "A2"

        # --- Дашборд по категориям ---
        ws_dash = self.wb.create_sheet("Дашборд_по_Категориям")
        ws_dash.cell(row=1, column=1, value="СВОДКА ПО КАТЕГОРИЯМ (FBS)").font = Font(bold=True, size=14, color="0F3460")
        dash_headers = ["Категория", "Кол-во SKU", "Общая Выручка", "Общая Прибыль", "Средняя Маржа %"]
        for col, h in enumerate(dash_headers, 1):
            ws_dash.cell(row=3, column=col, value=h)
        self._style_header(ws_dash, 3, len(dash_headers))

        # Получаем уникальные категории из входных данных
        categories = df['Категория'].dropna().unique().tolist()
        for i, cat in enumerate(categories):
            row = 4 + i
            ws_dash.cell(row=row, column=1, value=cat)
            # Используем SUMIFS по листу Расчет_FBS (категория в колонке B)
            ws_dash.cell(row=row, column=2, value=f'=COUNTIF(Расчет_FBS!B:B, A{row})')
            ws_dash.cell(row=row, column=3, value=f'=SUMIF(Расчет_FBS!B:B, A{row}, Расчет_FBS!C:C)').number_format = '#,##0.00 ₽'
            ws_dash.cell(row=row, column=4, value=f'=SUMIF(Расчет_FBS!B:B, A{row}, Расчет_FBS!K:K)').number_format = '#,##0.00 ₽'
            ws_dash.cell(row=row, column=5, value=f'=AVERAGEIF(Расчет_FBS!B:B, A{row}, Расчет_FBS!L:L)').number_format = '0.00%'
            for col in range(1, 6):
                ws_dash.cell(row=row, column=col).border = self.border
        ws_dash.column_dimensions['A'].width = 30
        for col in range(2, 6):
            ws_dash.column_dimensions[get_column_letter(col)].width = 18
        # Условное форматирование для средней маржи <0
        if categories:
            ws_dash.conditional_formatting.add(f'E4:E{3+len(categories)}',
                CellIsRule(operator='lessThan', formula=['0'],
                           fill=PatternFill(start_color='FFC7CE', end_color='FFC7CE', fill_type='solid')))

        # Удаляем стандартный лист
        if "Sheet" in self.wb.sheetnames:
            self.wb.remove(self.wb["Sheet"])

        self.wb.save(filepath)

# -----------------------------------------------------------------------------
# ОСНОВНОЙ UI STREAMLIT
# -----------------------------------------------------------------------------
def render_tariff_ui(tariff_mgr: TariffManager):
    st.header("⚙️ Настройка тарифа (Яндекс Маркет FBS)")
    tab1, tab2, tab3 = st.tabs(["📋 Текущий тариф", "📥 Загрузить CSV", "🔄 API Яндекс Маркет"])

    with tab1:
        st.subheader("Параметры тарифа")
        t = tariff_mgr.tariff
        cols = st.columns(4)
        cols[0].metric("Ставка комиссии", f"{t.commission_rate*100:.2f}%")
        cols[1].metric("Мин. комиссия", f"{t.min_commission:.2f} руб")
        cols[2].metric("Логистика база", f"{t.logistics_base:.2f} руб")
        cols[3].metric("Логистика за кг", f"{t.logistics_per_kg:.2f} руб/кг")
        cols2 = st.columns(3)
        cols2[0].metric("Хранение (л/день)", f"{t.storage_per_day_per_liter:.3f} руб")
        cols2[1].metric("Эквайринг", f"{t.acquiring_fee*100:.2f}%")
        cols2[2].metric("Возврат (резерв)", f"{t.return_fee*100:.2f}%")

        st.download_button(
            "⬇️ Скачать текущий тариф (CSV)",
            data=tariff_mgr.get_current_tariff_csv(),
            file_name="yandex_tariff.csv",
            mime="text/csv"
        )

    with tab2:
        st.subheader("Загрузить новый тариф из CSV")
        uploaded = st.file_uploader("CSV с разделителем ;", type=['csv'], key="tariff_csv")
        if uploaded and st.button("Применить CSV"):
            ok, msg = tariff_mgr.update_from_csv(uploaded.getvalue())
            if ok:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)

    with tab3:
        st.subheader("Загрузить тариф через API Яндекс Маркета")
        col1, col2 = st.columns(2)
        token = col1.text_input("OAuth токен", type="password")
        camp = col2.text_input("ID кампании", value="")
        if st.button("Получить и обновить"):
            if not token or not camp:
                st.error("Заполните оба поля")
            else:
                with st.spinner("Запрос к API..."):
                    ok, msg = tariff_mgr.fetch_from_yandex_api(token, int(camp))
                    if ok:
                        st.success(msg)
                        st.rerun()
                    else:
                        st.error(msg)

# -----------------------------------------------------------------------------
# КЭШИРОВАННАЯ ЗАГРУЗКА И ОБРАБОТКА ДАННЫХ
# -----------------------------------------------------------------------------
@st.cache_data
def load_and_process_data(file_content: bytes, file_name: str,
                          categories_db: AutoPartsCategoriesDB,
                          specific_costs: FBSSpecificCosts) -> pd.DataFrame:
    """Загружает файл, валидирует, добавляет спецрасходы, возвращает готовый DataFrame."""
    try:
        if file_name.endswith('.csv'):
            df = pd.read_csv(io.BytesIO(file_content), sep=';')
        else:
            df = pd.read_excel(io.BytesIO(file_content))
    except Exception as e:
        raise ValueError(f"Ошибка чтения файла: {e}")

    # Маппинг колонок
    col_map = {}
    for col in df.columns:
        cl = col.lower()
        if 'артикул' in cl or 'article' in cl:
            col_map[col] = 'Артикул'
        elif 'категория' in cl or 'category' in cl:
            col_map[col] = 'Категория'
        elif 'цена' in cl or 'price' in cl:
            col_map[col] = 'Цена'
        elif 'себестоимость' in cl or 'cost' in cl:
            col_map[col] = 'Себестоимость'
        elif 'вес' in cl or 'weight' in cl:
            col_map[col] = 'Вес_кг'
        elif 'длина' in cl or 'length' in cl:
            col_map[col] = 'Длина'
        elif 'ширина' in cl or 'width' in cl:
            col_map[col] = 'Ширина'
        elif 'высота' in cl or 'height' in cl:
            col_map[col] = 'Высота'
        elif 'объем' in cl or 'volume' in cl:
            col_map[col] = 'Объем_л'
    df = df.rename(columns=col_map)

    # Проверка наличия обязательных колонок
    required = ['Артикул', 'Цена', 'Себестоимость']
    for r in required:
        if r not in df.columns:
            raise ValueError(f"Не найдена колонка '{r}'. Убедитесь, что в файле есть нужные заголовки.")

    # Если нет категории, создаём пустую
    if 'Категория' not in df.columns:
        df['Категория'] = ''

    # Валидация габаритов
    df = FBSDimensionsValidator.validate_batch(df, categories_db)

    # Добавление спецрасходов
    df = specific_costs.calculate_batch(df)

    # Фильтрация: только положительная цена
    df = df[df['Цена'] > 0].copy()

    return df

# -----------------------------------------------------------------------------
# ВИЗУАЛИЗАЦИЯ (графики)
# -----------------------------------------------------------------------------
def render_plots(df: pd.DataFrame):
    if df.empty:
        st.info("Нет данных для визуализации")
        return

    # Метрики
    total_profit = df['Прибыль'].sum()
    avg_margin = df['Маржа_%'].mean() * 100
    count = len(df)
    col1, col2, col3 = st.columns(3)
    col1.metric("Количество SKU", f"{count:,}")
    col2.metric("Общая прибыль", f"{total_profit:,.2f} ₽", delta=None)
    col3.metric("Средняя маржа", f"{avg_margin:.1f}%", delta=None)

    # Два графика
    fig1 = px.histogram(df, x='Маржа_%', nbins=30, title='Распределение маржинальности',
                        labels={'Маржа_%':'Маржа, доля'}, color_discrete_sequence=['#0F3460'])
    fig1.add_vline(x=0, line_dash="dash", line_color="red")
    fig1.add_vline(x=0.15, line_dash="dash", line_color="green")

    # График прибыли по категориям (топ-10)
    cat_profit = df.groupby('Категория')['Прибыль'].sum().sort_values(ascending=False).head(10).reset_index()
    fig2 = px.bar(cat_profit, x='Категория', y='Прибыль', title='Топ-10 категорий по прибыли',
                  labels={'Прибыль':'Прибыль, руб'}, color='Прибыль', color_continuous_scale='Blues')

    st.plotly_chart(fig1, use_container_width=True)
    st.plotly_chart(fig2, use_container_width=True)

# -----------------------------------------------------------------------------
# ГЛАВНАЯ ФУНКЦИЯ РАСЧЁТА
# -----------------------------------------------------------------------------
def render_calculation_ui(tariff_mgr: TariffManager, categories_db: AutoPartsCategoriesDB,
                          specific_costs: FBSSpecificCosts):
    st.header("📊 Расчёт юнит-экономики (FBS)")

    # Настройка спецрасходов в сайдбаре (вынесем в sidebar)
    with st.sidebar.expander("🔧 Настройки спецрасходов", expanded=False):
        packaging = st.number_input("Упаковка FBS (руб)", value=specific_costs.packaging, step=1.0)
        chestny = st.number_input("Честный знак (руб)", value=specific_costs.chestny_znak, step=0.5)
        labeling = st.number_input("Маркировка (руб)", value=specific_costs.labeling, step=0.5)
        warranty = st.number_input("Гарантийный резерв (%)", value=specific_costs.warranty_reserve*100, step=0.1) / 100.0
        hazard = st.number_input("Надбавка за опасный груз (%)", value=specific_costs.hazard_surcharge*100, step=0.1) / 100.0
        fragile = st.number_input("Надбавка за хрупкость (%)", value=specific_costs.fragile_surcharge*100, step=0.1) / 100.0
        # Обновляем объект
        specific_costs.packaging = packaging
        specific_costs.chestny_znak = chestny
        specific_costs.labeling = labeling
        specific_costs.warranty_reserve = warranty
        specific_costs.hazard_surcharge = hazard
        specific_costs.fragile_surcharge = fragile

    uploaded_file = st.file_uploader("Загрузите каталог (CSV или Excel)", type=['csv', 'xlsx', 'xls'])

    if uploaded_file is not None:
        try:
            with st.spinner("Обработка данных..."):
                df = load_and_process_data(uploaded_file.getvalue(), uploaded_file.name,
                                           categories_db, specific_costs)

            if df.empty:
                st.warning("После фильтрации не осталось товаров с положительной ценой.")
                return

            st.success(f"✅ Обработано {len(df)} товаров")

            # ---- Добавляем колонки с результатами (для отображения) ----
            t = tariff_mgr.tariff
            df['Комиссия_руб'] = df.apply(lambda r: max(r['Цена'] * t.commission_rate, t.min_commission), axis=1)
            df['Логистика_руб'] = df.apply(lambda r: t.logistics_base + r['Оплач_вес'] * t.logistics_per_kg, axis=1)
            df['Хранение_руб'] = df['Объем_л'] * t.storage_per_day_per_liter * 30
            df['Эквайринг_руб'] = df['Цена'] * t.acquiring_fee
            df['Итого_расходы'] = (df['Себестоимость'] + df['Комиссия_руб'] + df['Логистика_руб'] +
                                   df['Хранение_руб'] + df['Эквайринг_руб'] + df['Спец_расходы_FBS'])
            df['Прибыль'] = df['Цена'] - df['Итого_расходы']
            df['Маржа_%'] = df.apply(lambda r: r['Прибыль'] / r['Цена'] if r['Цена'] > 0 else 0, axis=1)

            # ---- Отображение таблицы (сокращённой) ----
            with st.expander("📋 Предпросмотр данных (первые 100 строк)", expanded=False):
                st.dataframe(df.head(100), use_container_width=True)

            # ---- Графики и метрики ----
            render_plots(df)

            # ---- Экспорт в Excel ----
            if st.button("🚀 Сгенерировать Excel с формулами и дашбордом", type="primary"):
                with st.spinner("Создание Excel..."):
                    exporter = AdvancedExcelExporter(t, specific_costs)
                    temp_path = "unit_economics_fbs.xlsx"
                    exporter.export(df, temp_path)
                    with open(temp_path, "rb") as f:
                        st.download_button(
                            "⬇️ Скачать Excel",
                            data=f,
                            file_name="unit_economics_fbs.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    os.remove(temp_path)
                    st.success("✅ Excel файл готов!")

        except Exception as e:
            st.error(f"❌ Ошибка: {e}")
            st.code(traceback.format_exc())

# -----------------------------------------------------------------------------
# ЗАПУСК
# -----------------------------------------------------------------------------
def main():
    st.set_page_config(page_title="Юнит-экономика FBS", layout="wide", initial_sidebar_state="expanded")

    st.title(APP_NAME)
    st.caption(f"Версия {APP_VERSION} | Только Яндекс Маркет FBS")

    # Инициализация менеджеров
    tariff_mgr = TariffManager()
    categories_db = AutoPartsCategoriesDB()
    specific_costs = FBSSpecificCosts()

    # Меню
    menu = st.sidebar.radio("Навигация", ["⚙️ Тариф", "📊 Расчёт"], index=0)

    if menu == "⚙️ Тариф":
        render_tariff_ui(tariff_mgr)
    else:
        render_calculation_ui(tariff_mgr, categories_db, specific_costs)

    st.sidebar.markdown("---")
    st.sidebar.info("💡 **Подсказка**: загрузите файл с колонками: Артикул, Категория, Цена, Себестоимость, Вес_кг, Длина, Ширина, Высота (опционально).")

if __name__ == "__main__":
    main()
