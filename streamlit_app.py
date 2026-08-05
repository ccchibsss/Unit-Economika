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

# Excel (оптимизированный движок для больших объемов)
import xlsxwriter

# -----------------------------------------------------------------------------
# НАСТРОЙКА ЛОГГИРОВАНИЯ
# -----------------------------------------------------------------------------
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('UnitEconomicsFBS')

# -----------------------------------------------------------------------------
# КОНСТАНТЫ
# -----------------------------------------------------------------------------
APP_VERSION = "3.2.1"
APP_NAME = "🚗 Юнит-экономика (FBS Яндекс Маркет)"
BASE_DIR = Path(__file__).parent.resolve() if '__file__' in locals() else Path.cwd()
DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
TEMP_DIR = BASE_DIR / "temp"
for d in [DATA_DIR, CACHE_DIR, TEMP_DIR]:
    d.mkdir(parents=True, exist_ok=True)

TARIFF_JSON = DATA_DIR / "yandex_tariff.json"

# -----------------------------------------------------------------------------
# СПЕЦТАРИФЫ ДЛЯ АВТОЗАПЧАСТЕЙ (приоритет над базовыми)
# -----------------------------------------------------------------------------
SPECIAL_TARIFFS = {
    "шины": {"commission_rate": 0.12, "logistics_base": 90.0, "is_oversized": True},
    "аккумуляторы": {"commission_rate": 0.13, "logistics_base": 75.0, "hazardous": True},
    "двигатели": {"commission_rate": 0.11, "logistics_base": 120.0, "is_oversized": True},
    "кпп": {"commission_rate": 0.11, "logistics_base": 110.0, "is_oversized": True},
}

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
    commission_rate: float = 0.14
    min_commission: float = 45.0
    logistics_base: float = 45.0
    logistics_per_kg: float = 14.0
    storage_per_day_per_liter: float = 0.25
    acquiring_fee: float = 0.02
    return_fee: float = 0.02
    # Добавлены поля для корректной работы get_effective_tariff без TypeError
    is_oversized: bool = False
    hazardous: bool = False

    def to_dict(self) -> Dict:
        return asdict(self)

    @classmethod
    def from_dict(cls, data: Dict) -> 'YandexTariff':
        valid_keys = {k: v for k, v in data.items() if k in cls.__dataclass_fields__}
        return cls(**valid_keys)

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
        return cls()

    def get_effective_tariff(self, category: str) -> 'YandexTariff':
        cat_key = category.lower().strip()
        effective = asdict(self)
        for key, rules in SPECIAL_TARIFFS.items():
            if key in cat_key:
                effective.update(rules)
                break
        return YandexTariff(**effective)

# -----------------------------------------------------------------------------
# МЕНЕДЖЕР ТАРИФА
# -----------------------------------------------------------------------------
class TariffManager:
    def __init__(self):
        self.tariff = YandexTariff.load()

    def update_from_csv(self, file_content: bytes) -> Tuple[bool, str]:
        try:
            df = pd.read_csv(io.BytesIO(file_content), sep=';')
            required = ['commission_rate', 'logistics_base', 'logistics_per_kg', 'storage_per_day']
            if not all(c in df.columns for c in required):
                return False, f"Отсутствуют колонки: {required}"
            row = df.iloc[0]
            self.tariff.commission_rate = float(row.get('commission_rate', 0.14))
            self.tariff.min_commission = float(row.get('min_commission', 45.0))
            self.tariff.logistics_base = float(row.get('logistics_base', 45.0))
            self.tariff.logistics_per_kg = float(row.get('logistics_per_kg', 14.0))
            self.tariff.storage_per_day_per_liter = float(row.get('storage_per_day', 0.25))
            self.tariff.acquiring_fee = float(row.get('acquiring_fee', 0.02))
            self.tariff.return_fee = float(row.get('return_fee', 0.02))
            self.tariff.save()
            return True, "Тариф обновлён из CSV"
        except Exception as e:
            return False, f"Ошибка парсинга CSV: {e}"

    def fetch_from_yandex_api(self, oauth_token: str, campaign_id: int) -> Tuple[bool, str]:
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
            self.tariff.save()
            return True, f"Тариф обновлён из API (усреднено по {count} категориям)"
        except Exception as e:
            return False, f"Ошибка запроса: {e}"

    def get_tariff_dict(self) -> Dict:
        return self.tariff.to_dict()

    def get_current_tariff_csv(self) -> bytes:
        df = pd.DataFrame([self.tariff.to_dict()])
        return df.to_csv(index=False, sep=';', encoding='utf-8-sig').encode('utf-8-sig')

# -----------------------------------------------------------------------------
# СПРАВОЧНИК КАТЕГОРИЙ
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
# ВАЛИДАТОР ГАБАРИТОВ
# -----------------------------------------------------------------------------
class FBSDimensionsValidator:
    @staticmethod
    def normalize_dimension(value, unit_hint="cm"):
        unit = str(unit_hint).lower()
        if any(x in unit for x in ["mm", "мм"]):
            return value / 10.0
        if any(x in unit for x in ["m", "метр"]):
            return value * 100.0
        return float(value)

    @staticmethod
    def calculate_billable_weight(weight_kg: float, length_cm: float, width_cm: float, height_cm: float) -> float:
        if length_cm <= 0 or width_cm <= 0 or height_cm <= 0:
            return max(0.1, weight_kg)
        volumetric_weight = (length_cm * width_cm * height_cm) / 5000.0
        return max(weight_kg, volumetric_weight)

    @staticmethod
    def validate_batch(df: pd.DataFrame, categories_db: AutoPartsCategoriesDB) -> pd.DataFrame:
        df = df.copy()
        for col in ['Длина', 'Ширина', 'Высота']:
            if col in df.columns:
                df[col] = df[col].apply(lambda x: FBSDimensionsValidator.normalize_dimension(safe_float(x)))
        
        if 'Вес_кг' in df.columns:
            df['Вес_кг'] = df['Вес_кг'].apply(lambda x: safe_float(x) if safe_float(x) < 100 else safe_float(x)/1000)
            
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
                
        if all(col in df.columns for col in ['Длина', 'Ширина', 'Высота']):
            df['Оплач_вес'] = df.apply(
                lambda r: FBSDimensionsValidator.calculate_billable_weight(
                    safe_float(r['Вес_кг']), safe_float(r['Длина']),
                    safe_float(r['Ширина']), safe_float(r['Высота'])
                ), axis=1
            )
        else:
            df['Оплач_вес'] = df['Вес_кг']
        return df

# -----------------------------------------------------------------------------
# РАСЧЁТ СПЕЦИФИЧЕСКИХ РАСХОДОВ
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
        df['Маркировка'] = self.chestny_znak + self.labeling
        df['Гарант_резерв'] = df['Цена'] * self.warranty_reserve
        
        if 'is_hazardous' not in df.columns: df['is_hazardous'] = False
        if 'is_fragile' not in df.columns: df['is_fragile'] = False
            
        df['Надбавка_опасный'] = np.where(df['is_hazardous'], df['Цена'] * self.hazard_surcharge, 0.0)
        df['Надбавка_хрупкий'] = np.where(df['is_fragile'], df['Цена'] * self.fragile_surcharge, 0.0)
        df['Спец_расходы_FBS'] = (df['Упаковка_FBS'] + df['Маркировка'] +
                                  df['Гарант_резерв'] + df['Надбавка_опасный'] +
                                  df['Надбавка_хрупкий'])
        return df

# -----------------------------------------------------------------------------
# РАСЧЁТ МЕТРИК В PANDAS (Критически важно для работы дашборда Streamlit)
# -----------------------------------------------------------------------------
def calculate_metrics_in_pandas(df: pd.DataFrame, tariff_manager: TariffManager) -> pd.DataFrame:
    df = df.copy()
    
    def calc_row(row):
        eff_tariff = tariff_manager.tariff.get_effective_tariff(str(row.get('Категория', '')))
        
        comm = max(row['Цена'] * eff_tariff.commission_rate, eff_tariff.min_commission)
        logist = eff_tariff.logistics_base + row.get('Оплач_вес', row['Вес_кг']) * eff_tariff.logistics_per_kg
        storage = row.get('Объем_л', 0) * eff_tariff.storage_per_day_per_liter * row.get('Оборачиваемость_дней', 30)
        acquiring = row['Цена'] * eff_tariff.acquiring_fee
        
        total_expenses = row['Себестоимость'] + comm + logist + storage + acquiring + row.get('Спец_расходы_FBS', 0)
        profit = row['Цена'] - total_expenses
        margin = (profit / row['Цена']) if row['Цена'] > 0 else 0.0
        
        cat_lower = str(row.get('Категория', '')).lower()
        is_special = any(k in cat_lower for k in ["шины", "аккумуляторы", "двигатели", "двигатель", "кпп"])
        reason = ""
        if is_special:
            if "шины" in cat_lower: reason = "Крупногабаритный"
            elif "аккумуляторы" in cat_lower: reason = "Опасный груз"
            elif "двигатель" in cat_lower or "двигатели" in cat_lower or "кпп" in cat_lower: reason = "Крупногабаритный/тяжёлый"

        return pd.Series({
            'Комиссия_руб': comm, 'Логистика_руб': logist, 'Хранение_руб': storage,
            'Эквайринг_руб': acquiring, 'Итого_расходы': total_expenses,
            'Прибыль': profit, 'Маржа_%': margin, 'Спецтариф_применён': is_special, 
            'Причина_спецтарифа': reason
        })

    metrics_df = df.apply(calc_row, axis=1)
    return pd.concat([df, metrics_df], axis=1)

# -----------------------------------------------------------------------------
# ЭКСПОРТ В EXCEL
# -----------------------------------------------------------------------------
class AdvancedExcelExporter:
    def __init__(self, tariff: YandexTariff, specific_costs: FBSSpecificCosts):
        self.tariff = tariff
        self.specific_costs = specific_costs

    def export(self, df: pd.DataFrame, filepath: str):
        with pd.ExcelWriter(filepath, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            fmt_header = workbook.add_format({'bold': True, 'bg_color': '#0F3460', 'font_color': '#FFFFFF', 'border': 1, 'align': 'center', 'valign': 'vcenter'})
            fmt_money = workbook.add_format({'num_format': '#,##0.00', 'border': 1})
            fmt_percent = workbook.add_format({'num_format': '0.00%', 'border': 1})
            fmt_red = workbook.add_format({'bg_color': '#FFC7CE', 'border': 1})
            fmt_green = workbook.add_format({'bg_color': '#C6EFCE', 'border': 1})
            fmt_formula_bg = workbook.add_format({'bg_color': '#DCE6F1', 'border': 1})

            # 1. Лист Тариф
            df_tariff = pd.DataFrame([self.tariff.to_dict()])
            df_tariff.to_excel(writer, sheet_name='Тариф', index=False)
            ws_tariff = writer.sheets['Тариф']
            for col_num, value in enumerate(df_tariff.columns.values):
                ws_tariff.write(0, col_num, value, fmt_header)
            workbook.define_name('TariffRow', '=Тариф!$A$2:$G$2')

            # 2. Лист Входные_Данные
            in_cols = ['Артикул', 'Категория', 'Цена', 'Себестоимость', 'Вес_кг', 
                       'Длина', 'Ширина', 'Высота', 'Объем_л', 'Оплач_вес', 'Спец_расходы_FBS', 'Оборачиваемость_дней']
            if 'Оборачиваемость_дней' not in df.columns:
                df['Оборачиваемость_дней'] = 30
            
            existing_in_cols = [c for c in in_cols if c in df.columns]
            df_in = df[existing_in_cols].copy()
            df_in.to_excel(writer, sheet_name='Входные_Данные', index=False)
            ws_in = writer.sheets['Входные_Данные']
            ws_in.freeze_panes(1, 0)
            ws_in.set_column('A:M', 15)

            # 3. Лист Расчет_FBS
            ws_calc = workbook.add_worksheet('Расчет_FBS')
            last_row = len(df) + 1
            table_range = f'A1:N{last_row}'

            # Исправлены формулы: русские функции (МАКС, ЕСЛИ) и разделитель точка с запятой (;)
            table_columns = [
                {'header': 'Артикул', 'formula': '=Входные_Данные!A2'},
                {'header': 'Категория', 'formula': '=Входные_Данные!B2'},
                {'header': 'Цена', 'formula': '=Входные_Данные!C2'},
                {'header': 'Себестоимость', 'formula': '=Входные_Данные!D2'},
                {'header': 'Комиссия_руб', 'formula': '=МАКС(C2*Тариф!$B$2; Тариф!$C$2)'},
                {'header': 'Логистика_руб', 'formula': '=Тариф!$D$2 + Входные_Данные!J2*Тариф!$E$2'},
                {'header': 'Хранение_руб', 'formula': '=Входные_Данные!I2*Тариф!$F$2*Входные_Данные!M2'},
                {'header': 'Эквайринг_руб', 'formula': '=C2*Тариф!$G$2'},
                {'header': 'Спец_расходы_FBS', 'formula': '=Входные_Данные!K2'},
                {'header': 'Итого_расходы', 'formula': '=D2+E2+F2+G2+H2+I2'},
                {'header': 'Прибыль', 'formula': '=C2-J2'},
                {'header': 'Маржа_%', 'formula': '=ЕСЛИ(C2>0; K2/C2; 0)'},
                {'header': 'Спецтариф_применён', 
                 'formula': ('=ЕСЛИ(ИЛИ(ЕЧИСЛО(ПОИСК("шины";B2)); ЕЧИСЛО(ПОИСК("аккумуляторы";B2)); '
                             'ЕЧИСЛО(ПОИСК("двигатель";B2)); ЕЧИСЛО(ПОИСК("кпп";B2))); ИСТИНА; ЛОЖЬ)')},
                {'header': 'Причина_спецтарифа',
                 'formula': ('=ЕСЛИ(Спецтариф_применён; '
                             'ЕСЛИ(ЕЧИСЛО(ПОИСК("шины";B2)); "Крупногабаритный"; '
                             'ЕСЛИ(ЕЧИСЛО(ПОИСК("аккумуляторы";B2)); "Опасный груз"; '
                             'ЕСЛИ(ЕЧИСЛО(ПОИСК("двигатель";B2)); "Крупногабаритный/тяжёлый"; '
                             'ЕСЛИ(ЕЧИСЛО(ПОИСК("кпп";B2)); "Крупногабаритный/тяжёлый"; ""))))); "")')}
            ]

            ws_calc.add_table(table_range, {'columns': table_columns, 'style': 'Table Style Medium 2', 'banded_rows': True})

            ws_calc.set_column('C:D', 12, fmt_money)
            ws_calc.set_column('E:I', 14, fmt_formula_bg)
            ws_calc.set_column('J:K', 14, fmt_money)
            ws_calc.set_column('L:L', 12, fmt_percent)
            ws_calc.set_column('M:N', 25)
            ws_calc.freeze_panes(1, 0)

            ws_calc.conditional_format(f'L2:L{last_row}', {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt_red})
            ws_calc.conditional_format(f'L2:L{last_row}', {'type': 'cell', 'criteria': '>=', 'value': 0.15, 'format': fmt_green})
            ws_calc.conditional_format(f'K2:K{last_row}', {'type': 'cell', 'criteria': '<', 'value': 0, 'format': fmt_red})

            # 4. Дашборд по категориям
            ws_dash = workbook.add_worksheet('Дашборд_по_Категориям')
            ws_dash.write('A1', 'СВОДКА ПО КАТЕГОРИЯМ (FBS)', workbook.add_format({'bold': True, 'size': 14, 'color': '#0F3460'}))
            dash_headers = ["Категория", "Кол-во SKU", "Общая Выручка", "Общая Прибыль", "Средняя Маржа %", "Спецтариф_доля_%"]
            ws_dash.write_row(2, 0, dash_headers, fmt_header)
            
            categories = df['Категория'].dropna().unique().tolist()
            for i, cat in enumerate(categories):
                row = 3 + i
                ws_dash.write(row, 0, cat)
                ws_dash.write_formula(row, 1, f'=СЧЁТЕСЛИ(Расчет_FBS!B:B; A{row+1})')
                ws_dash.write_formula(row, 2, f'=СУММЕСЛИ(Расчет_FBS!B:B; A{row+1}; Расчет_FBS!C:C)')
                ws_dash.write_formula(row, 3, f'=СУММЕСЛИ(Расчет_FBS!B:B; A{row+1}; Расчет_FBS!K:K)')
                ws_dash.write_formula(row, 4, f'=СРЗНАЧЕСЛИ(Расчет_FBS!B:B; A{row+1}; Расчет_FBS!L:L)')
                ws_dash.write_formula(row, 5, f'=(СЧЁТЕСЛИМН(Расчет_FBS!B:B; A{row+1}; Расчет_FBS!M:M; ИСТИНА) / СЧЁТЗ(Расчет_FBS!B:B)) * 100')
                
            ws_dash.set_column('A:A', 30)
            ws_dash.set_column('B:B', 12)
            ws_dash.set_column('C:D', 16, fmt_money)
            ws_dash.set_column('E:E', 16, fmt_percent)
            ws_dash.set_column('F:F', 16, fmt_percent)

            # 5. Лист Легенда
            ws_legend = workbook.add_worksheet('Легенда')
            legend_data = [
                ["Колонка", "Описание", "Источник данных / логика"],
                ["Артикул", "Уникальный идентификатор товара", "Ввод пользователя"],
                ["Категория", "Группа товаров для подбора габаритов и спецтарифов", "Ввод пользователя / справочник"],
                ["Цена", "Розничная цена продажи на Маркете", "Ввод пользователя"],
                ["Себестоимость", "Закупочная цена + прямые затраты", "Ввод пользователя"],
                ["Комиссия_руб", "МАКС(Цена*Ставка; Мин_комиссия). Спецтарифы применяются автоматически", "Тариф + спецправила"],
                ["Логистика_руб", "База + Оплач_вес*Ставка_за_кг", "Тариф + габариты"],
                ["Хранение_руб", "Объем_л*Ставка*Оборачиваемость_дней", "Тариф + входные данные"],
                ["Эквайринг_руб", "Цена*Ставка эквайринга", "Тариф"],
                ["Спец_расходы_FBS", "Упаковка, маркировка, резервы, надбавки за хрупкость/опасность", "Настройки FBSSpecificCosts"],
                ["Итого_расходы", "Себестоимость + Сумма всех комиссий и сборов маркета", "Формула"],
                ["Прибыль", "Цена - Итого_расходов", "Формула"],
                ["Маржа_%", "Прибыль / Цена", "Формула"],
                ["Спецтариф_применён", "Флаг применения спецтарифа по ключевым словам категории", "Логика SPECIAL_TARIFFS"],
                ["Причина_спецтарифа", "Текстовое пояснение причины спецтарифа", "Логика SPECIAL_TARIFFS"]
            ]
            for r_idx, row_data in enumerate(legend_data):
                for c_idx, val in enumerate(row_data):
                    if r_idx == 0:
                        ws_legend.write(r_idx, c_idx, val, fmt_header)
                    else:
                        ws_legend.write(r_idx, c_idx, val)
            ws_legend.set_column('A:C', 40)

# -----------------------------------------------------------------------------
# ОСНОВНОЙ ИНТЕРФЕЙС STREAMLIT
# -----------------------------------------------------------------------------
def main():
    st.set_page_config(page_title=APP_NAME, layout="wide")
    st.title(APP_NAME)
    st.write(f"Версия: {APP_VERSION}")

    tariff_manager = TariffManager()
    categories_db = AutoPartsCategoriesDB()
    specific_costs = FBSSpecificCosts()

    with st.sidebar:
        st.header("Управление тарифами")
        uploaded_file = st.file_uploader("Загрузить тариф из CSV (; разделитель)", type=["csv"])
        if uploaded_file:
            success, msg = tariff_manager.update_from_csv(uploaded_file.read())
            if success:
                st.success(msg)
            else:
                st.error(msg)

        st.divider()
        oauth_token = st.text_input("OAuth токен Яндекс Маркета", type="password")
        campaign_id = st.number_input("Campaign ID", min_value=1, value=123456)
        if st.button("Обновить тарифы из API"):
            if oauth_token:
                success, msg = tariff_manager.fetch_from_yandex_api(oauth_token, campaign_id)
                if success:
                    st.success(msg)
                else:
                    st.error(msg)
            else:
                st.warning("Введите OAuth токен")

        st.download_button(
            label="Скачать текущий тариф (CSV)",
            data=tariff_manager.get_current_tariff_csv(),
            file_name="current_tariff.csv"
        )

    st.divider()
    st.subheader("Загрузка входных данных (Excel/CSV)")
    input_file = st.file_uploader("Загрузите файл с товарами (колонки: Артикул, Категория, Цена, Себестоимость, Вес_кг, Длина, Ширина, Высота, Объем_л)", type=["csv", "xlsx"])

    if input_file:
        try:
            if input_file.name.endswith(".csv"):
                df_raw = pd.read_csv(input_file, sep=";")
            else:
                df_raw = pd.read_excel(input_file)

            required_cols = ['Артикул', 'Категория', 'Цена', 'Себестоимость']
            missing = [c for c in required_cols if c not in df_raw.columns]
            if missing:
                st.error(f"Отсутствуют обязательные колонки: {missing}")
            else:
                df_validated = FBSDimensionsValidator.validate_batch(df_raw, categories_db)
                df_costs = specific_costs.calculate_batch(df_validated)
                
                # Расчет метрик в Pandas для предотвращения KeyError в дашборде
                df_final = calculate_metrics_in_pandas(df_costs, tariff_manager)

                st.success("Данные загружены, валидированы и рассчитаны")
                with st.expander("Предварительный просмотр данных"):
                    st.dataframe(df_final)

                export_filename = f"unit_economy_fbs_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                exporter = AdvancedExcelExporter(tariff_manager.tariff, specific_costs)
                exporter.export(df_final, export_filename)

                with open(export_filename, "rb") as f:
                    st.download_button(
                        label="📥 Скачать Excel с живыми формулами и дашбордом",
                        data=f,
                        file_name=export_filename
                    )

                st.subheader("Ключевые метрики")
                total_revenue = df_final['Цена'].sum()
                total_expenses = df_final['Итого_расходы'].sum()
                total_profit = df_final['Прибыль'].sum()
                avg_margin = (total_profit / total_revenue) * 100 if total_revenue > 0 else 0

                metrics = {
                    "Общая выручка": f"{money_round(total_revenue):,} ₽",
                    "Итого расходы": f"{money_round(total_expenses):,} ₽",
                    "Чистая прибыль": f"{money_round(total_profit):,} ₽",
                    "Средняя маржа": f"{money_round(avg_margin, 1)}%"
                }
                cols = st.columns(4)
                for col, (k, v) in zip(cols, metrics.items()):
                    col.metric(label=k, value=v)

                if 'Категория' in df_final.columns and 'Маржа_%' in df_final.columns:
                    df_plot = df_final.groupby('Категория')['Маржа_%'].mean().reset_index()
                    fig = px.bar(df_plot, x='Категория', y='Маржа_%', title='Средняя маржа по категориям',
                                 labels={'Маржа_%': 'Маржа (%)', 'Категория': 'Категория'},
                                 color='Маржа_%', color_continuous_scale='RdYlGn')
                    st.plotly_chart(fig, use_container_width=True)

        except Exception as e:
            logger.exception(e)
            st.error(f"Ошибка обработки файла: {e}")

if __name__ == "__main__":
    main()

# === КОНЕЦ КОДА ===
