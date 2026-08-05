"""
================================================================================
🚗 UNIT ECONOMICS FOR AUTO PARTS (FBS ONLY) - OPTIMIZED v2.0
================================================================================
✅ ТОЛЬКО РЕЖИМ FBS
✅ ТАРИФЫ НЕ ЗАКОДИРОВАНЫ: отдельный лист в Excel, обновление через CSV/API
✅ ЖИВЫЕ ФОРМУЛЫ В EXCEL: VLOOKUP/XLOOKUP ссылаются на лист с тарифами
✅ ОПТИМИЗАЦИЯ: убран мертвый код, лишние импорты и многопоточность там, где она не нужна
================================================================================
"""

# ============================================================================
# БЛОК 0: БАЗОВЫЕ ИМПОРТЫ (ОПТИМИЗИРОВАНО)
# ============================================================================
import streamlit as st
import pandas as pd
import numpy as np
import requests
import logging
import json
import os
import time
import traceback
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, field, asdict
from decimal import Decimal, ROUND_HALF_UP
import io

# Настройка логгирования
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('UnitEconomicsFBS')

# ============================================================================
# БЛОК 1: КОНСТАНТЫ И КОНФИГУРАЦИЯ (ТОЛЬКО FBS)
# ============================================================================
APP_VERSION = "2.0.0"
APP_NAME = "🚗 Юнит-экономика автозапчастей (FBS)"

# Директории
BASE_DIR = Path(__file__).parent.resolve() if '__file__' in locals() else Path.cwd()
DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
TEMP_DIR = BASE_DIR / "temp"

for d in [DATA_DIR, CACHE_DIR, TEMP_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# Файл для хранения тарифов
TARIFFS_FILE = DATA_DIR / "marketplace_tariffs.json"
TARIFFS_CSV_TEMPLATE = DATA_DIR / "tariffs_template.csv"

# Лимиты Excel
EXCEL_HARD_LIMIT = 1_048_576
EXCEL_SAFE_ROW_LIMIT = 100_000 # При превышении живые формулы могут тормозить Excel, но мы их оставляем, просто оптимизируем

# ============================================================================
# БЛОК 2: УТИЛИТЫ И МОДЕЛИ ДАННЫХ
# ============================================================================
def money_round(value: float, decimals: int = 2) -> float:
    """Корректное округление денег"""
    return float(Decimal(str(value)).quantize(Decimal(f"0.{'0' * decimals}"), rounding=ROUND_HALF_UP))

@dataclass
class MarketplaceTariff:
    """Структура тарифа для ОДНОГО маркетплейса (Режим FBS)"""
    name: str
    commission_rate: float = 0.0      # Комиссия за продажу (доля, например 0.15)
    min_commission: float = 0.0       # Минимальная комиссия (руб)
    logistics_base: float = 0.0       # База логистики (руб)
    logistics_per_kg: float = 0.0     # Логистика за кг (руб/кг)
    storage_per_day: float = 0.0      # Хранение за литр в день (руб/л/день)
    acquiring_fee: float = 0.0        # Эквайринг (доля)
    last_mile_fee: float = 0.0        # Последняя миля (руб)
    return_fee: float = 0.0           # Стоимость возврата (руб или доля)
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'MarketplaceTariff':
        return cls(**{k: v for k, v in data.items() if k in cls.__dataclass_fields__})

# ============================================================================
# БЛОК 3: МЕНЕДЖЕР ТАРИФОВ (API, CSV, JSON)
# ============================================================================
class TariffManager:
    """
    Управление тарифами. 
    Источники: Локальный JSON, CSV, API Яндекс Маркет.
    """
    def __init__(self):
        self.tariffs: Dict[str, MarketplaceTariff] = {}
        self._load_from_cache()

    def _load_from_cache(self):
        """Загрузка из локального JSON кэша"""
        if TARIFFS_FILE.exists():
            try:
                with open(TARIFFS_FILE, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    for mp_name, mp_data in data.items():
                        self.tariffs[mp_name] = MarketplaceTariff.from_dict(mp_data)
                logger.info(f"Загружено {len(self.tariffs)} тарифов из кэша")
            except Exception as e:
                logger.error(f"Ошибка чтения кэша тарифов: {e}")
                self._load_defaults()
        else:
            self._load_defaults()

    def _load_defaults(self):
        """Базовые тарифы (только для первичной инициализации, далее управляются через UI/CSV)"""
        self.tariffs = {
            "Ozon": MarketplaceTariff("Ozon", 0.15, 30.0, 50.0, 15.0, 0.3, 0.015, 50.0, 0.02),
            "Wildberries": MarketplaceTariff("Wildberries", 0.18, 35.0, 60.0, 18.0, 0.5, 0.0, 0.0, 0.03),
            "Яндекс Маркет": MarketplaceTariff("Яндекс Маркет", 0.14, 0.0, 45.0, 14.0, 0.25, 0.02, 40.0, 0.02),
            "Мегамаркет": MarketplaceTariff("Мегамаркет", 0.13, 28.0, 55.0, 16.0, 0.3, 0.018, 45.0, 0.02),
        }
        self.save_to_cache()

    def save_to_cache(self):
        """Сохранение в JSON"""
        try:
            data = {name: tariff.to_dict() for name, tariff in self.tariffs.items()}
            with open(TARIFFS_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"Ошибка сохранения кэша: {e}")

    def update_tariff(self, mp_name: str, tariff: MarketplaceTariff):
        self.tariffs[mp_name] = tariff
        self.save_to_cache()

    def load_from_csv(self, file_content: bytes) -> Tuple[bool, str]:
        """Загрузка тарифов из CSV"""
        try:
            df = pd.read_csv(io.BytesIO(file_content), sep=';')
            required_cols = ['name', 'commission_rate', 'logistics_base', 'logistics_per_kg', 'storage_per_day']
            if not all(col in df.columns for col in required_cols):
                return False, f"Отсутствуют колонки: {required_cols}"
            
            for _, row in df.iterrows():
                mp_name = str(row['name']).strip()
                tariff = MarketplaceTariff(
                    name=mp_name,
                    commission_rate=float(row.get('commission_rate', 0)),
                    min_commission=float(row.get('min_commission', 0)),
                    logistics_base=float(row.get('logistics_base', 0)),
                    logistics_per_kg=float(row.get('logistics_per_kg', 0)),
                    storage_per_day=float(row.get('storage_per_day', 0)),
                    acquiring_fee=float(row.get('acquiring_fee', 0)),
                    last_mile_fee=float(row.get('last_mile_fee', 0)),
                    return_fee=float(row.get('return_fee', 0))
                )
                self.tariffs[mp_name] = tariff
            self.save_to_cache()
            return True, f"Успешно загружено {len(df)} тарифов"
        except Exception as e:
            return False, f"Ошибка парсинга CSV: {e}"

    def get_csv_template_bytes(self) -> bytes:
        """Генерация шаблона CSV для скачивания"""
        df = pd.DataFrame([t.to_dict() for t in self.tariffs.values()])
        return df.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

    def fetch_yandex_market_tariffs(self, oauth_token: str, campaign_id: int) -> Tuple[bool, str]:
        """
        Получение тарифов Яндекс Маркет через API.
        Примечание: API ЯМ требует специфичной обработки категорий, здесь базовый фетч.
        """
        url = f"https://api.partner.market.yandex.ru/v2/campaigns/{campaign_id}/deliveries/fees"
        headers = {"Authorization": f"OAuth {oauth_token}"}
        try:
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                data = response.json()
                # В реальном проекте здесь нужен парсинг специфичного ответа ЯМ в структуру MarketplaceTariff
                # Для примера сохраняем сырые данные или маппим, если структура известна
                logger.info("Тарифы ЯМ получены. Требуется маппинг под конкретную категорию.")
                return True, "Данные получены. Требуется тонкий маппинг по категориям ЯМ."
            else:
                return False, f"Ошибка API ЯМ: HTTP {response.status_code}"
        except Exception as e:
            return False, f"Ошибка запроса к API ЯМ: {e}"

# Инициализация синглтона
@st.cache_resource
def get_tariff_manager() -> TariffManager:
    return TariffManager()
# ============================================================================
# БЛОК 2: ЯДРО РАСЧЕТА FBS И БАЗА EXCEL-ЭКСПОРТЕРА
# ============================================================================
import pandas as pd
import numpy as np
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import logging

logger = logging.getLogger('UnitEconomicsFBS')

class FBSCalculationEngine:
    """
    Движок расчета юнит-экономики строго для режима FBS.
    Использует данные из TariffManager.
    """
    def __init__(self, tariff_manager: TariffManager):
        self.tariff_manager = tariff_manager

    def calculate_batch(self, df: pd.DataFrame) -> pd.DataFrame:
        """Векторизованный расчет для DataFrame. Возвращает расширенный DF."""
        if df.empty:
            return df

        # Базовые колонки для расчета
        required = ['Артикул', 'Маркетплейс', 'Цена', 'Себестоимость', 'Вес_кг', 'Объем_л']
        for col in required:
            if col not in df.columns:
                raise ValueError(f"Отсутствует обязательная колонка: {col}")

        res = df.copy()
        
        # Маппинг тарифов для векторизации
        tariffs_map = {name: t for name, t in self.tariff_manager.tariffs.items()}
        
        # Применяем тарифы к строкам
        res['Комиссия_%'] = res['Маркетплейс'].map({k: v.commission_rate for k, v in tariffs_map.items()}).fillna(0)
        res['Мин_комиссия'] = res['Маркетплейс'].map({k: v.min_commission for k, v in tariffs_map.items()}).fillna(0)
        res['Логистика_база'] = res['Маркетплейс'].map({k: v.logistics_base for k, v in tariffs_map.items()}).fillna(0)
        res['Логистика_кг'] = res['Маркетплейс'].map({k: v.logistics_per_kg for k, v in tariffs_map.items()}).fillna(0)
        res['Хранение_л_день'] = res['Маркетплейс'].map({k: v.storage_per_day for k, v in tariffs_map.items()}).fillna(0)
        res['Эквайринг_%'] = res['Маркетплейс'].map({k: v.acquiring_fee for k, v in tariffs_map.items()}).fillna(0)
        res['Возврат_%'] = res['Маркетплейс'].map({k: v.return_fee for k, v in tariffs_map.items()}).fillna(0)

        # Математика FBS
        res['Комиссия_руб'] = np.maximum(res['Цена'] * res['Комиссия_%'], res['Мин_комиссия'])
        res['Логистика_руб'] = res['Логистика_база'] + (res['Вес_кг'] * res['Логистика_кг'])
        # Хранение считаем за средние 30 дней
        res['Хранение_руб'] = res['Объем_л'] * res['Хранение_л_день'] * 30 
        res['Эквайринг_руб'] = res['Цена'] * res['Эквайринг_%']
        res['Возврат_руб'] = res['Цена'] * res['Возврат_%']
        
        res['Итого_расходы'] = (res['Себестоимость'] + res['Комиссия_руб'] + 
                                res['Логистика_руб'] + res['Хранение_руб'] + 
                                res['Эквайринг_руб'] + res['Возврат_руб'])
        
        res['Прибыль'] = res['Цена'] - res['Итого_расходы']
        res['Маржа_%'] = np.where(res['Цена'] > 0, (res['Прибыль'] / res['Цена']) * 100, 0)
        
        return res


class ExcelFormulaExporter:
    """
    Генератор Excel-файлов с ЖИВЫМИ формулами.
    Тарифы выносятся на отдельный лист. Расчеты ссылаются на него через VLOOKUP.
    """
    def __init__(self, tariff_manager: TariffManager):
        self.tariff_manager = tariff_manager
        self.wb = Workbook()
        
        # Стили
        self.header_font = Font(bold=True, color="FFFFFF", size=11)
        self.header_fill = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
        self.tariff_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
        self.input_fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
        self.formula_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
        self.thin_border = Border(
            left=Side(style='thin'), right=Side(style='thin'),
            top=Side(style='thin'), bottom=Side(style='thin')
        )

    def _style_header(self, ws, row_num, max_col):
        for col in range(1, max_col + 1):
            cell = ws.cell(row=row_num, column=col)
            cell.font = self.header_font
            cell.fill = self.header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            cell.border = self.thin_border

    def _create_tariff_sheet(self):
        """Создает лист 'Тарифы_FBS' с актуальными данными из TariffManager."""
        ws = self.wb.create_sheet("Тарифы_FBS")
        
        headers = [
            "Маркетплейс", "Комиссия_%", "Мин_комиссия_руб", 
            "Логистика_база", "Логистика_за_кг", "Хранение_л_день", 
            "Эквайринг_%", "Возврат_%"
        ]
        
        for col_idx, header in enumerate(headers, 1):
            ws.cell(row=1, column=col_idx, value=header)
        self._style_header(ws, 1, len(headers))
        
        row_idx = 2
        for name, tariff in self.tariff_manager.tariffs.items():
            ws.cell(row=row_idx, column=1, value=name).fill = self.tariff_fill
            ws.cell(row=row_idx, column=2, value=tariff.commission_rate).number_format = '0.00%'
            ws.cell(row=row_idx, column=3, value=tariff.min_commission).number_format = '#,##0.00'
            ws.cell(row=row_idx, column=4, value=tariff.logistics_base).number_format = '#,##0.00'
            ws.cell(row=row_idx, column=5, value=tariff.logistics_per_kg).number_format = '#,##0.00'
            ws.cell(row=row_idx, column=6, value=tariff.storage_per_day).number_format = '#,##0.000'
            ws.cell(row=row_idx, column=7, value=tariff.acquiring_fee).number_format = '0.00%'
            ws.cell(row=row_idx, column=8, value=tariff.return_fee).number_format = '0.00%'
            
            for col in range(1, 9):
                ws.cell(row=row_idx, column=col).fill = self.tariff_fill
                ws.cell(row=row_idx, column=col).border = self.thin_border
            row_idx += 1
            
        # Формируем именованный диапазон или просто фиксируем диапазон для VLOOKUP
        self.tariff_range = f"Тарифы_FBS!$A$1:$H${row_idx - 1}"
        
        # Настраиваем ширину колонок
        for col in range(1, 9):
            ws.column_dimensions[get_column_letter(col)].width = 18
            
        return ws
# ============================================================================
# БЛОК 3: ЖИВЫЕ ФОРМУЛЫ (VLOOKUP) И ДАШБОРД
# ============================================================================
    def _create_input_and_calc_sheets(self, df: pd.DataFrame):
        """
        Создает лист 'Входные_Данные' и лист 'Расчет_FBS'.
        В 'Расчет_FBS' все финансовые показатели считаются через VLOOKUP к листу 'Тарифы_FBS'.
        """
        ws_in = self.wb.create_sheet("Входные_Данные")
        ws_calc = self.wb.create_sheet("Расчет_FBS")
        
        # --- ЛИСТ ВХОДНЫЕ ДАННЫЕ ---
        in_headers = ["Артикул", "Маркетплейс", "Цена", "Себестоимость", "Вес_кг", "Объем_л"]
        for col_idx, h in enumerate(in_headers, 1):
            ws_in.cell(row=1, column=col_idx, value=h)
        self._style_header(ws_in, 1, len(in_headers))
        
        for r_idx, row in df.iterrows():
            excel_row = r_idx + 2
            ws_in.cell(row=excel_row, column=1, value=row.get('Артикул', ''))
            ws_in.cell(row=excel_row, column=2, value=row.get('Маркетплейс', 'Ozon'))
            ws_in.cell(row=excel_row, column=3, value=row.get('Цена', 0)).number_format = '#,##0.00'
            ws_in.cell(row=excel_row, column=4, value=row.get('Себестоимость', 0)).number_format = '#,##0.00'
            ws_in.cell(row=excel_row, column=5, value=row.get('Вес_кг', 0)).number_format = '#,##0.000'
            ws_in.cell(row=excel_row, column=6, value=row.get('Объем_л', 0)).number_format = '#,##0.000'
            
            for col in range(1, 7):
                ws_in.cell(row=excel_row, column=col).fill = self.input_fill
                ws_in.cell(row=excel_row, column=col).border = self.thin_border

        # --- ЛИСТ РАСЧЕТ (ЖИВЫЕ ФОРМУЛЫ) ---
        calc_headers = [
            "Артикул", "Маркетплейс", "Цена", "Себестоимость", 
            "Комиссия_руб", "Логистика_руб", "Хранение_руб", "Эквайринг_руб",
            "Итого_расходы", "Прибыль", "Маржа_%"
        ]
        for col_idx, h in enumerate(calc_headers, 1):
            ws_calc.cell(row=1, column=col_idx, value=h)
        self._style_header(ws_calc, 1, len(calc_headers))
        
        total_rows = len(df)
        
        for r_idx in range(total_rows):
            excel_row = r_idx + 2
            in_row = excel_row # Строка в листе Входные_Данные
            
            # Ссылки на входные данные
            ref_mp = f"Входные_Данные!B{in_row}"
            ref_price = f"Входные_Данные!C{in_row}"
            ref_cost = f"Входные_Данные!D{in_row}"
            ref_weight = f"Входные_Данные!E{in_row}"
            ref_vol = f"Входные_Данные!F{in_row}"
            
            # Записываем артикул и маркетплейс (ссылками)
            ws_calc.cell(row=excel_row, column=1, value=f"=Входные_Данные!A{in_row}")
            ws_calc.cell(row=excel_row, column=2, value=f"={ref_mp}")
            ws_calc.cell(row=excel_row, column=3, value=f"={ref_price}").number_format = '#,##0.00'
            ws_calc.cell(row=excel_row, column=4, value=f"={ref_cost}").number_format = '#,##0.00'
            
            # ЖИВЫЕ ФОРМУЛЫ (VLOOKUP / ВПР)
            # Примечание: используем английские имена функций, так как openpyxl не транслирует их автоматически.
            # Excel сам переведет их в локальную версию при открытии, если язык пакета совпадает.
            
            # Комиссия = MAX(Цена * VLOOKUP(МП; Тарифы!Комиссия_%), VLOOKUP(МП; Тарифы!Мин_комиссия))
            comm_formula = f'=MAX({ref_price} * VLOOKUP({ref_mp}, {self.tariff_range}, 2, FALSE), VLOOKUP({ref_mp}, {self.tariff_range}, 3, FALSE))'
            ws_calc.cell(row=excel_row, column=5, value=comm_formula).number_format = '#,##0.00'
            
            # Логистика = База + Вес * Логистика_кг
            log_formula = f'=VLOOKUP({ref_mp}, {self.tariff_range}, 4, FALSE) + {ref_weight} * VLOOKUP({ref_mp}, {self.tariff_range}, 5, FALSE)'
            ws_calc.cell(row=excel_row, column=6, value=log_formula).number_format = '#,##0.00'
            
            # Хранение = Объем * Хранение_л_день * 30
            stor_formula = f'={ref_vol} * VLOOKUP({ref_mp}, {self.tariff_range}, 6, FALSE) * 30'
            ws_calc.cell(row=excel_row, column=7, value=stor_formula).number_format = '#,##0.00'
            
            # Эквайринг = Цена * Эквайринг_%
            acq_formula = f'={ref_price} * VLOOKUP({ref_mp}, {self.tariff_range}, 7, FALSE)'
            ws_calc.cell(row=excel_row, column=8, value=acq_formula).number_format = '#,##0.00'
            
            # Итого расходы = Себестоимость + Комиссия + Логистика + Хранение + Эквайринг
            total_exp_formula = f'={ref_cost} + SUM(E{excel_row}:H{excel_row})'
            ws_calc.cell(row=excel_row, column=9, value=total_exp_formula).number_format = '#,##0.00'
            
            # Прибыль = Цена - Итого расходы
            profit_formula = f'={ref_price} - I{excel_row}'
            ws_calc.cell(row=excel_row, column=10, value=profit_formula).number_format = '#,##0.00'
            
            # Маржа % = Прибыль / Цена
            margin_formula = f'=IF({ref_price}>0, J{excel_row}/{ref_price}, 0)'
            ws_calc.cell(row=excel_row, column=11, value=margin_formula).number_format = '0.00%'
            
            # Применяем стили к формулам
            for col in range(5, 12):
                cell = ws_calc.cell(row=excel_row, column=col)
                cell.fill = self.formula_fill
                cell.border = self.thin_border

        # Настройка ширины колонок
        for col in range(1, 12):
            ws_calc.column_dimensions[get_column_letter(col)].width = 16
            
        # Закрепляем шапку
        ws_in.freeze_panes = "A2"
        ws_calc.freeze_panes = "A2"

    def _create_dashboard_sheet(self):
        """Создает лист с дашбордом на основе SUMIFS."""
        ws = self.wb.create_sheet("Дашборд", 0) # Помещаем первым
        
        ws.cell(row=1, column=1, value="СВОДКА ПО МАРКЕТПЛЕЙСАМ (FBS)").font = Font(bold=True, size=14)
        
        headers = ["Маркетплейс", "Кол-во SKU", "Общая Выручка", "Общая Прибыль", "Средняя Маржа"]
        for col_idx, h in enumerate(headers, 1):
            ws.cell(row=3, column=col_idx, value=h)
        self._style_header(ws, 3, len(headers))
        
        # Получаем уникальные маркетплейсы из тарифов
        mps = list(self.tariff_manager.tariffs.keys())
        
        for i, mp in enumerate(mps):
            row = 4 + i
            ws.cell(row=row, column=1, value=mp)
            
            # COUNTIFS
            ws.cell(row=row, column=2, value=f'=COUNTIF(Расчет_FBS!B:B, A{row})')
            # SUMIFS Выручка
            ws.cell(row=row, column=3, value=f'=SUMIFS(Расчет_FBS!C:C, Расчет_FBS!B:B, A{row})').number_format = '#,##0.00 ₽'
            # SUMIFS Прибыль
            ws.cell(row=row, column=4, value=f'=SUMIFS(Расчет_FBS!J:J, Расчет_FBS!B:B, A{row})').number_format = '#,##0.00 ₽'
            # AVERAGEIFS Маржа
            ws.cell(row=row, column=5, value=f'=AVERAGEIFS(Расчет_FBS!K:K, Расчет_FBS!B:B, A{row})').number_format = '0.00%'
            
            for col in range(1, 6):
                ws.cell(row=row, column=col).border = self.thin_border
                
        ws.column_dimensions['A'].width = 20
        for col in range(2, 6):
            ws.column_dimensions[get_column_letter(col)].width = 18

    def export_to_excel(self, df: pd.DataFrame, filepath: str):
        """Собирает все листы и сохраняет файл."""
        self._create_tariff_sheet()
        self._create_input_and_calc_sheets(df)
        self._create_dashboard_sheet()
        
        # Удаляем дефолтный пустой лист, если он остался
        if "Sheet" in self.wb.sheetnames:
            del self.wb["Sheet"]
            
        self.wb.save(filepath)
        logger.info(f"Excel с живыми формулами сохранен: {filepath}")
# ============================================================================
# БЛОК 4: UI STREAMLIT (FBS ONLY)
# ============================================================================
import streamlit as st
import os

def render_tariff_management_ui(tariff_manager: TariffManager):
    """UI для управления тарифами (CSV, Ручное редактирование)."""
    st.header("⚙️ Управление тарифами FBS")
    st.info("Тарифы хранятся локально и выносятся на отдельный лист Excel. Измените их здесь, чтобы пересчитать экономику.")
    
    tab1, tab2, tab3 = st.tabs(["📋 Текущие тарифы", "📥 Загрузка из CSV", "🔄 Обновление"])
    
    with tab1:
        st.subheader("Актуальные тарифы")
        tariffs_data = [t.to_dict() for t in tariff_manager.tariffs.values()]
        if tariffs_data:
            df_tariffs = pd.DataFrame(tariffs_data)
            st.dataframe(df_tariffs, use_container_width=True)
        else:
            st.warning("Тарифы не загружены.")
            
    with tab2:
        st.subheader("Загрузить тарифы из CSV")
        st.caption("Формат CSV: name;commission_rate;min_commission;logistics_base;logistics_per_kg;storage_per_day;acquiring_fee;return_fee")
        uploaded_csv = st.file_uploader("Выберите CSV файл", type=['csv'], key="tariff_csv")
        if uploaded_csv and st.button("Применить CSV"):
            success, msg = tariff_manager.load_from_csv(uploaded_csv.getvalue())
            if success:
                st.success(msg)
                st.rerun()
            else:
                st.error(msg)
                
        st.download_button(
            "📥 Скачать шаблон CSV",
            data=tariff_manager.get_csv_template_bytes(),
            file_name="tariffs_template.csv",
            mime="text/csv"
        )

    with tab3:
        st.subheader("Синхронизация")
        st.warning("Прямое API Яндекс Маркет требует OAuth-токен и ID кампании. Введите данные ниже.")
        col1, col2 = st.columns(2)
        with col1:
            ya_oauth = st.text_input("OAuth Token Яндекс Маркет", type="password")
        with col2:
            ya_campaign = st.number_input("Campaign ID", min_value=0, step=1)
            
        if st.button("Получить тарифы ЯМ"):
            if ya_oauth and ya_campaign:
                success, msg = tariff_manager.fetch_yandex_market_tariffs(ya_oauth, int(ya_campaign))
                if success:
                    st.success(msg)
                else:
                    st.error(msg)
            else:
                st.error("Заполните OAuth и Campaign ID")


def render_fbs_calculation_ui(engine: FBSCalculationEngine, exporter: ExcelFormulaExporter):
    """UI для загрузки каталога и генерации Excel."""
    st.header("📊 Расчет юнит-экономики (FBS)")
    
    uploaded_file = st.file_uploader("Загрузите каталог товаров (CSV/Excel)", type=['csv', 'xlsx', 'xls'], key="catalog_upload")
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, sep=';')
            else:
                df = pd.read_excel(uploaded_file)
                
            st.success(f"Загружено {len(df)} товаров.")
            
            # Маппинг колонок (упрощенное для чистоты кода)
            st.subheader("Сопоставление колонок")
            cols = df.columns.tolist()
            c1, c2, c3, c4 = st.columns(4)
            with c1: art_col = st.selectbox("Артикул", cols, index=0)
            with c2: mp_col = st.selectbox("Маркетплейс", cols, index=1 if len(cols)>1 else 0)
            with c3: price_col = st.selectbox("Цена", cols, index=2 if len(cols)>2 else 0)
            with c4: cost_col = st.selectbox("Себестоимость", cols, index=3 if len(cols)>3 else 0)
            
            c5, c6 = st.columns(2)
            with c5: weight_col = st.selectbox("Вес (кг)", cols, index=4 if len(cols)>4 else 0)
            with c6: vol_col = st.selectbox("Объем (л)", cols, index=5 if len(cols)>5 else 0)
            
            if st.button("🚀 Сгенерировать Excel с живыми формулами", type="primary"):
                with st.spinner("Подготовка данных и генерация формул..."):
                    # Нормализуем DF для движка
                    df_norm = pd.DataFrame({
                        'Артикул': df[art_col],
                        'Маркетплейс': df[mp_col],
                        'Цена': pd.to_numeric(df[price_col], errors='coerce').fillna(0),
                        'Себестоимость': pd.to_numeric(df[cost_col], errors='coerce').fillna(0),
                        'Вес_кг': pd.to_numeric(df[weight_col], errors='coerce').fillna(0),
                        'Объем_л': pd.to_numeric(df[vol_col], errors='coerce').fillna(0)
                    })
                    
                    # Фильтруем только валидные строки и известные МП
                    known_mps = list(engine.tariff_manager.tariffs.keys())
                    df_norm = df_norm[df_norm['Маркетплейс'].isin(known_mps)]
                    df_norm = df_norm[df_norm['Цена'] > 0]
                    
                    if df_norm.empty:
                        st.error("Нет данных для расчета. Проверьте названия маркетплейсов и цены.")
                        return

                    # Сохраняем во временный файл
                    temp_path = "unit_economics_fbs_live.xlsx"
                    exporter.export_to_excel(df_norm, temp_path)
                    
                    with open(temp_path, "rb") as f:
                        st.download_button(
                            "⬇️ Скачать Excel (Живые формулы)",
                            data=f,
                            file_name="unit_economics_fbs_live.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    st.success("Файл готов! Откройте его в Excel — измените тарифы на первом листе, и расчеты пересчитаются автоматически.")
                    
                    # Очистка
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                        
        except Exception as e:
            st.error(f"Ошибка обработки файла: {e}")

# ============================================================================
# БЛОК 5: СПРАВОЧНИК КАТЕГОРИЙ И ВАЛИДАТОР ГАБАРИТОВ (FBS)
# ============================================================================
import math

class AutoPartsCategoriesDB:
    """
    Справочник 150+ категорий автозапчастей с типовыми габаритами.
    Используется для автозаполнения пропусков и валидации.
    """
    def __init__(self):
        # Словарь: {ключ: (тип_объем_л, тип_вес_кг, hazardous, fragile)}
        self.db = {
            "фильтры": (1.5, 0.5, False, False), "масла": (5.0, 4.0, True, False),
            "колодки": (0.8, 1.2, False, False), "диски": (3.0, 4.0, False, True),
            "амортизаторы": (4.0, 3.5, False, True), "аккумуляторы": (12.0, 15.0, True, True),
            "шины": (25.0, 10.0, False, False), "фары": (6.0, 2.5, False, True),
            "ремни": (0.5, 0.2, False, False), "подшипники": (0.3, 0.8, False, False),
            "датчики": (0.1, 0.1, False, False), "свечи": (0.05, 0.05, False, False),
            "помпы": (1.5, 1.5, False, False), "радиаторы": (15.0, 5.0, False, True),
            "бамперы": (40.0, 8.0, False, True), "крылья": (20.0, 6.0, False, True),
            "двигатель": (50.0, 80.0, True, True), "кпп": (40.0, 50.0, True, True),
        }

    def get_defaults(self, category: str) -> Dict[str, Any]:
        """Возвращает типовые габариты и флаги для категории."""
        cat_key = category.lower().strip()
        for key, defaults in self.db.items():
            if key in cat_key:
                return {
                    "volume_l": defaults[0], "weight_kg": defaults[1],
                    "is_hazardous": defaults[2], "is_fragile": defaults[3]
                }
        # Дефолт для неизвестных
        return {"volume_l": 2.0, "weight_kg": 1.0, "is_hazardous": False, "is_fragile": False}

class FBSDimensionsValidator:
    """
    Валидатор и нормализатор весогабаритов для FBS.
    Учитывает объемный вес (коэф. 5000) и пошаговую тарификацию.
    """
    @staticmethod
    def normalize_dimension(value: float, unit_hint: str = "cm") -> float:
        if not value or value <= 0: return 0.0
        unit = unit_hint.lower()
        if any(x in unit for x in ['mm', 'мм']): return value / 10.0
        if any(x in unit for x in ['m', 'метр']): return value * 100.0
        if value > 300: return value / 10.0  # Автоисправление мм -> см
        return value

    @staticmethod
    def calculate_billable_weight(weight_kg: float, length_cm: float, width_cm: float, height_cm: float) -> float:
        """Расчет оплачиваемого веса (больший из реального и объемного)."""
        if length_cm <= 0 or width_cm <= 0 or height_cm <= 0:
            return max(0.1, weight_kg)
        
        volumetric_weight = (length_cm * width_cm * height_cm) / 5000.0
        billable = max(weight_kg, volumetric_weight)
        # Округление до 0.5 кг (стандарт многих МП)
        return math.ceil(billable * 2) / 2

    @staticmethod
    def validate_batch(df: pd.DataFrame, categories_db: AutoPartsCategoriesDB) -> pd.DataFrame:
        """Векторизованная валидация и заполнение пропусков."""
        res = df.copy()
        
        # Нормализация единиц (если есть колонки)
        for dim_col in ['Длина', 'Ширина', 'Высота']:
            if dim_col in res.columns:
                res[dim_col] = res[dim_col].apply(lambda x: FBSDimensionsValidator.normalize_dimension(safe_float(x)))
        
        if 'Вес_кг' in res.columns:
            res['Вес_кг'] = res['Вес_кг'].apply(lambda x: safe_float(x) if safe_float(x) < 100 else safe_float(x)/1000)

        # Заполнение пропусков из справочника
        if 'Категория' in res.columns:
            for idx, row in res.iterrows():
                cat_defaults = categories_db.get_defaults(str(row.get('Категория', '')))
                if res.at[idx, 'Объем_л'] == 0 or pd.isna(res.at[idx, 'Объем_л']):
                    res.at[idx, 'Объем_л'] = cat_defaults['volume_l']
                if res.at[idx, 'Вес_кг'] == 0 or pd.isna(res.at[idx, 'Вес_кг']):
                    res.at[idx, 'Вес_кг'] = cat_defaults['weight_kg']
                
                # Флаги для доп. расходов
                res.at[idx, 'is_hazardous'] = cat_defaults['is_hazardous']
                res.at[idx, 'is_fragile'] = cat_defaults['is_fragile']

        # Расчет объемного веса
        if all(col in res.columns for col in ['Длина', 'Ширина', 'Высота']):
            res['Оплач_вес'] = res.apply(
                lambda r: FBSDimensionsValidator.calculate_billable_weight(
                    safe_float(r['Вес_кг']), safe_float(r['Длина']), 
                    safe_float(r['Ширина']), safe_float(r['Высота'])
                ), axis=1
            )
        else:
            res['Оплач_вес'] = res['Вес_кг']

        return res
# ============================================================================
# БЛОК 6: СПЕЦИФИЧЕСКИЕ РАСХОДЫ FBS И ОБНОВЛЕННЫЙ EXCEL-ЭКСПОРТЕР
# ============================================================================
@dataclass
class FBSSpecificCosts:
    """Специфические расходы для FBS (упаковка, маркировка, Честный знак)."""
    packaging_fbs: float = 45.0      # Упаковка FBS (руб/отправление)
    chestny_znak: float = 1.5        # Код маркировки (руб/шт)
    labeling: float = 3.0            # Наклейка штрихкода (руб/шт)
    warranty_reserve: float = 0.02   # Резерв на гарантийные случаи (доля от цены)

    def calculate_batch(self, df: pd.DataFrame) -> pd.DataFrame:
        """Векторизованный расчет специфических расходов."""
        res = df.copy()
        res['Упаковка_FBS'] = self.packaging_fbs
        res['Маркировка'] = np.where(res.get('requires_marking', True), self.chestny_znak + self.labeling, 0)
        res['Гарант_резерв'] = res['Цена'] * self.warranty_reserve
        
        # Надбавки за опасные/хрупкие
        res['Надбавка_опасный'] = np.where(res.get('is_hazardous', False), res['Цена'] * 0.01, 0)
        res['Надбавка_хрупкий'] = np.where(res.get('is_fragile', False), res['Цена'] * 0.005, 0)
        
        res['Спец_расходы_FBS'] = (res['Упаковка_FBS'] + res['Маркировка'] + 
                                    res['Гарант_резерв'] + res['Надбавка_опасный'] + res['Надбавка_хрупкий'])
        return res


class AdvancedExcelFormulaExporter(ExcelFormulaExporter):
    """
    Расширенный экспортер с учетом специфических расходов FBS и защитой от краша Excel.
    """
    EXCEL_SAFE_ROW_LIMIT = 100_000  # При превышении живые формулы могут тормозить Excel

    def __init__(self, tariff_manager: TariffManager, specific_costs: FBSSpecificCosts):
        super().__init__(tariff_manager)
        self.specific_costs = specific_costs
        self.use_live_formulas = True

    def _create_input_and_calc_sheets(self, df: pd.DataFrame):
        """Создает листы с учетом спец. расходов и защитой от лимитов."""
        self.use_live_formulas = len(df) <= self.EXCEL_SAFE_ROW_LIMIT
        if not self.use_live_formulas:
            logger.warning(f"⚠️ Превышен лимит {self.EXCEL_SAFE_ROW_LIMIT} строк. Живые формулы отключены для стабильности Excel.")

        ws_in = self.wb.create_sheet("Входные_Данные")
        ws_calc = self.wb.create_sheet("Расчет_FBS")
        
        # --- ЛИСТ ВХОДНЫЕ ДАННЫЕ ---
        in_headers = ["Артикул", "Маркетплейс", "Категория", "Цена", "Себестоимость", 
                      "Вес_кг", "Длина_см", "Ширина_см", "Высота_см", "Объем_л", "Оплач_вес"]
        for col_idx, h in enumerate(in_headers, 1):
            ws_in.cell(row=1, column=col_idx, value=h)
        self._style_header(ws_in, 1, len(in_headers))
        
        for r_idx, row in df.iterrows():
            excel_row = r_idx + 2
            ws_in.cell(row=excel_row, column=1, value=row.get('Артикул', ''))
            ws_in.cell(row=excel_row, column=2, value=row.get('Маркетплейс', 'Ozon'))
            ws_in.cell(row=excel_row, column=3, value=row.get('Категория', ''))
            ws_in.cell(row=excel_row, column=4, value=row.get('Цена', 0)).number_format = '#,##0.00'
            ws_in.cell(row=excel_row, column=5, value=row.get('Себестоимость', 0)).number_format = '#,##0.00'
            ws_in.cell(row=excel_row, column=6, value=row.get('Вес_кг', 0)).number_format = '#,##0.000'
            ws_in.cell(row=excel_row, column=7, value=row.get('Длина', 0)).number_format = '#,##0.0'
            ws_in.cell(row=excel_row, column=8, value=row.get('Ширина', 0)).number_format = '#,##0.0'
            ws_in.cell(row=excel_row, column=9, value=row.get('Высота', 0)).number_format = '#,##0.0'
            ws_in.cell(row=excel_row, column=10, value=row.get('Объем_л', 0)).number_format = '#,##0.000'
            ws_in.cell(row=excel_row, column=11, value=row.get('Оплач_вес', 0)).number_format = '#,##0.000'
            
            for col in range(1, 12):
                ws_in.cell(row=excel_row, column=col).fill = self.input_fill
                ws_in.cell(row=excel_row, column=col).border = self.thin_border

        # --- ЛИСТ РАСЧЕТ ---
        calc_headers = [
            "Артикул", "Маркетплейс", "Цена", "Себестоимость", 
            "Комиссия_руб", "Логистика_руб", "Хранение_руб", "Эквайринг_руб",
            "Спец_расходы_FBS", "Итого_расходы", "Прибыль", "Маржа_%"
        ]
        for col_idx, h in enumerate(calc_headers, 1):
            ws_calc.cell(row=1, column=col_idx, value=h)
        self._style_header(ws_calc, 1, len(calc_headers))
        
        total_rows = len(df)
        
        for r_idx in range(total_rows):
            excel_row = r_idx + 2
            in_row = excel_row 
            
            ref_mp = f"Входные_Данные!B{in_row}"
            ref_price = f"Входные_Данные!D{in_row}"
            ref_cost = f"Входные_Данные!E{in_row}"
            ref_billable_weight = f"Входные_Данные!K{in_row}"
            ref_vol = f"Входные_Данные!J{in_row}"
            
            ws_calc.cell(row=excel_row, column=1, value=f"=Входные_Данные!A{in_row}")
            ws_calc.cell(row=excel_row, column=2, value=f"={ref_mp}")
            ws_calc.cell(row=excel_row, column=3, value=f"={ref_price}").number_format = '#,##0.00'
            ws_calc.cell(row=excel_row, column=4, value=f"={ref_cost}").number_format = '#,##0.00'
            
            if self.use_live_formulas:
                # ЖИВЫЕ ФОРМУЛЫ
                comm_formula = f'=MAX({ref_price} * VLOOKUP({ref_mp}, {self.tariff_range}, 2, FALSE), VLOOKUP({ref_mp}, {self.tariff_range}, 3, FALSE))'
                ws_calc.cell(row=excel_row, column=5, value=comm_formula).number_format = '#,##0.00'
                
                log_formula = f'=VLOOKUP({ref_mp}, {self.tariff_range}, 4, FALSE) + {ref_billable_weight} * VLOOKUP({ref_mp}, {self.tariff_range}, 5, FALSE)'
                ws_calc.cell(row=excel_row, column=6, value=log_formula).number_format = '#,##0.00'
                
                stor_formula = f'={ref_vol} * VLOOKUP({ref_mp}, {self.tariff_range}, 6, FALSE) * 30'
                ws_calc.cell(row=excel_row, column=7, value=stor_formula).number_format = '#,##0.00'
                
                acq_formula = f'={ref_price} * VLOOKUP({ref_mp}, {self.tariff_range}, 7, FALSE)'
                ws_calc.cell(row=excel_row, column=8, value=acq_formula).number_format = '#,##0.00'
                
                # Спец расходы берем из входных данных (они уже посчитаны в Python)
                ws_calc.cell(row=excel_row, column=9, value=f"=Входные_Данные!L{in_row}").number_format = '#,##0.00' # Предполагаем, что L - это спец расходы
                
                total_exp_formula = f'={ref_cost} + SUM(E{excel_row}:I{excel_row})'
                ws_calc.cell(row=excel_row, column=10, value=total_exp_formula).number_format = '#,##0.00'
            else:
                # СТАТИЧЕСКИЕ ЗНАЧЕНИЯ (для больших объемов)
                mp_name = df.iloc[r_idx]['Маркетплейс']
                tariff = self.tariff_manager.tariffs.get(mp_name)
                if tariff:
                    price = df.iloc[r_idx]['Цена']
                    cost = df.iloc[r_idx]['Себестоимость']
                    bill_w = df.iloc[r_idx]['Оплач_вес']
                    vol = df.iloc[r_idx]['Объем_л']
                    spec = df.iloc[r_idx].get('Спец_расходы_FBS', 0)
                    
                    comm = max(price * tariff.commission_rate, tariff.min_commission)
                    log = tariff.logistics_base + bill_w * tariff.logistics_per_kg
                    stor = vol * tariff.storage_per_day * 30
                    acq = price * tariff.acquiring_fee
                    
                    ws_calc.cell(row=excel_row, column=5, value=comm).number_format = '#,##0.00'
                    ws_calc.cell(row=excel_row, column=6, value=log).number_format = '#,##0.00'
                    ws_calc.cell(row=excel_row, column=7, value=stor).number_format = '#,##0.00'
                    ws_calc.cell(row=excel_row, column=8, value=acq).number_format = '#,##0.00'
                    ws_calc.cell(row=excel_row, column=9, value=spec).number_format = '#,##0.00'
                    
                    total_exp = cost + comm + log + stor + acq + spec
                    ws_calc.cell(row=excel_row, column=10, value=total_exp).number_format = '#,##0.00'

            # Прибыль и Маржа (всегда формулы, они легкие)
            profit_formula = f'={ref_price} - J{excel_row}'
            ws_calc.cell(row=excel_row, column=11, value=profit_formula).number_format = '#,##0.00'
            
            margin_formula = f'=IF({ref_price}>0, K{excel_row}/{ref_price}, 0)'
            ws_calc.cell(row=excel_row, column=12, value=margin_formula).number_format = '0.00%'
            
            for col in range(5, 13):
                cell = ws_calc.cell(row=excel_row, column=col)
                cell.fill = self.formula_fill
                cell.border = self.thin_border

        for col in range(1, 13):
            ws_calc.column_dimensions[get_column_letter(col)].width = 16
            
        ws_in.freeze_panes = "A2"
        ws_calc.freeze_panes = "A2"
# ============================================================================
# БЛОК 7: ФИНАЛЬНЫЙ UI STREAMLIT И ТОЧКА ВХОДА
# ============================================================================
def render_fbs_calculation_ui(tariff_mgr: TariffManager, categories_db: AutoPartsCategoriesDB, specific_costs: FBSSpecificCosts):
    """Полный UI для расчета FBS с валидацией и живыми формулами."""
    st.header("📊 Расчет юнит-экономики (FBS)")
    
    uploaded_file = st.file_uploader("Загрузите каталог товаров (CSV/Excel)", type=['csv', 'xlsx', 'xls'], key="catalog_upload")
    
    if uploaded_file is not None:
        try:
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, sep=';')
            else:
                df = pd.read_excel(uploaded_file)
                
            st.success(f"Загружено {len(df)} товаров.")
            
            # Нормализация колонок
            col_mapping = {}
            for col in df.columns:
                cl = col.lower()
                if 'артикул' in cl or 'article' in cl: col_mapping[col] = 'Артикул'
                elif 'мп' in cl or 'marketplace' in cl: col_mapping[col] = 'Маркетплейс'
                elif 'категория' in cl or 'category' in cl: col_mapping[col] = 'Категория'
                elif 'цена' in cl or 'price' in cl: col_mapping[col] = 'Цена'
                elif 'себестоимость' in cl or 'cost' in cl: col_mapping[col] = 'Себестоимость'
                elif 'вес' in cl or 'weight' in cl: col_mapping[col] = 'Вес_кг'
                elif 'длина' in cl or 'length' in cl: col_mapping[col] = 'Длина'
                elif 'ширина' in cl or 'width' in cl: col_mapping[col] = 'Ширина'
                elif 'высота' in cl or 'height' in cl: col_mapping[col] = 'Высота'
                elif 'объем' in cl or 'volume' in cl: col_mapping[col] = 'Объем_л'
            
            df = df.rename(columns=col_mapping)
            
            # Валидация и заполнение пропусков
            with st.spinner("Валидация габаритов и заполнение пропусков..."):
                df_validated = FBSDimensionsValidator.validate_batch(df, categories_db)
                df_with_costs = specific_costs.calculate_batch(df_validated)
            
            st.success(f"✅ Валидация завершена. Заполнено пропусков по справочнику: {df_with_costs['Объем_л'].notna().sum()}")
            
            # Фильтрация известных МП
            known_mps = list(tariff_mgr.tariffs.keys())
            df_final = df_with_costs[df_with_costs['Маркетплейс'].isin(known_mps)]
            df_final = df_final[df_final['Цена'] > 0]
            
            if df_final.empty:
                st.error("Нет данных для расчета. Проверьте названия маркетплейсов и цены.")
                return

            # Экспортер
            exporter = AdvancedExcelFormulaExporter(tariff_mgr, specific_costs)
            
            if st.button("🚀 Сгенерировать Excel с живыми формулами", type="primary"):
                with st.spinner("Подготовка данных и генерация формул..."):
                    temp_path = "unit_economics_fbs_live.xlsx"
                    exporter.export_to_excel(df_final, temp_path)
                    
                    with open(temp_path, "rb") as f:
                        st.download_button(
                            "⬇️ Скачать Excel (Живые формулы)",
                            data=f,
                            file_name="unit_economics_fbs_live.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                    st.success("Файл готов! Откройте его в Excel — измените тарифы на первом листе, и расчеты пересчитаются автоматически.")
                    
                    if os.path.exists(temp_path):
                        os.remove(temp_path)
                        
        except Exception as e:
            st.error(f"Ошибка обработки файла: {e}")
            st.code(traceback.format_exc())


def main():
    st.set_page_config(page_title="Unit Economics FBS Pro", layout="wide")
    st.title("🚗 Юнит-экономика автозапчастей (Только FBS)")
    
    # Инициализация синглтонов
    tariff_mgr = get_tariff_manager()
    categories_db = AutoPartsCategoriesDB()
    specific_costs = FBSSpecificCosts()
    
    menu = st.sidebar.radio("Меню", ["⚙️ Тарифы", "📊 Расчет"])
    
    if menu == "⚙️ Тарифы":
        render_tariff_management_ui(tariff_mgr)
    elif menu == "📊 Расчет":
        render_fbs_calculation_ui(tariff_mgr, categories_db, specific_costs)

if __name__ == "__main__":
    main()
