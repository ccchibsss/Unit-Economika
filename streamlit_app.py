#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
============================================================================
🚀 ULTIMATE UNIT ECONOMICS FOR AUTO PARTS v15.0 - ENTERPRISE SYNC EDITION
============================================================================
"""

import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime
import io
import json
import requests
import logging
import warnings
import re
from enum import Enum
from decimal import Decimal, ROUND_HALF_UP
from typing import Dict, List, Tuple, Optional

warnings.filterwarnings('ignore')
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger('YandexMarketHybridEconomics')

OPENPYXL_AVAILABLE = False
try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule
    OPENPYXL_AVAILABLE = True
except ImportError:
    pass

APP_VERSION = "15.0.0"
APP_NAME = "FBS Unit Economics & Price Strategy Manager"

# ============================================================================
# БЛОК 0: СЛУЖЕБНЫЕ УТИЛИТЫ ТОЧНЫХ РАСЧЕТОВ И ИСПРАВЛЕНИЯ КОДИРОВОК (MOJIBAKE)
# ============================================================================
def money_round(value: float) -> float:
    """Точное финансовое округление до 2 знаков после запятой (ROUND_HALF_UP)."""
    if np.isnan(value) or np.isinf(value):
        return 0.0
    return float(Decimal(str(value)).quantize(Decimal("0.00"), rounding=ROUND_HALF_UP))

def fix_double_utf8(text: str) -> str:
    """Исправление кракозябр (двойного UTF-8 кодирования) при импорте каталогов."""
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

# ============================================================================
# БЛОК 1: КОНФИГУРАЦИИ И СТРУКТУРЫ ДАННЫХ
# ============================================================================
class TaxSystem(Enum):
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
            if item.label == label: return item
        return cls.USN_6

class Tariff:
    def __init__(self, category: str, commission_rate: float = 0.12, min_commission: float = 35.0,
                 magma_base: float = 30.0, magma_per_kg: float = 15.0, acquiring_fee: float = 0.018,
                 return_fee: float = 0.05, source: str = "Справочник оферты"):
        self.category = category
        self.commission_rate = commission_rate
        self.min_commission = min_commission
        self.magma_base = magma_base
        self.magma_per_kg = magma_per_kg
        self.acquiring_fee = acquiring_fee
        self.return_fee = return_fee
        self.source = source

# ============================================================================
# БЛОК 2: МОДУЛЬ ГИБРИДНОГО ПОЛУЧЕНИЯ ТАРИФОВ (API + AI DEEPSEEK + LOCAL)
# ============================================================================
class HybridTariffManager:
    DEFAULTS = {
        'default': Tariff('default', 0.12, 35, 30, 15, 0.018, 0.05, "Локальная база"),
        'автозапчасти': Tariff('автозапчасти', 0.10, 30, 30, 15, 0.018, 0.06, "Оферта автозапчасти"),
        'электроника': Tariff('электроника', 0.08, 30, 30, 15, 0.015, 0.04, "Оферта электроники"),
        'одежда': Tariff('одежда', 0.15, 25, 25, 12, 0.018, 0.07, "Оферта одежда")
    }

    def __init__(self):
        self.tariffs = dict(self.DEFAULTS)

    def fetch_from_yandex_market_api(self, oauth_token: str, campaign_id: str) -> bool:
        url = f"[link removed]{campaign_id}/deliveries/fees"
        headers = {"Authorization": f"OAuth {oauth_token}", "Content-Type": "application/json"}
        try:
            response = requests.get(url, headers=headers, timeout=10)
            if response.status_code == 200:
                data = response.json()
                if "fees" in data:
                    for fee in data.get("fees", []):
                        type_f = fee.get("deliveryServiceType", "default")
                        self.tariffs[type_f.lower()] = Tariff(
                            category=type_f,
                            commission_rate=float(fee.get("commissionRate", 12)) / 100,
                            source="Яндекс.Маркет API Live"
                        )
                return True
            return False
        except Exception as e:
            logger.error(f"Ошибка вызова API Яндекс Маркета: {e}")
            return False

    def fetch_from_deepseek_ai(self, api_key: str, category_name: str) -> Tariff:
        if not api_key: return self.tariffs['default']
        url = "[link removed]"
        headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}
        prompt = f"Проанализируй категорию товара '{category_name}' на Яндекс Маркет FBS. Выдай JSON с ключами commission_rate (доля от 0 до 1), return_rate (доля от 0 до 1). Только чистый JSON."
        try:
            payload = {"model": "deepseek-chat", "messages": [{"role": "user", "content": prompt}], "response_format": {"type": "json_object"}}
            res = requests.post(url, json=payload, headers=headers, timeout=10)
            if res.status_code == 200:
                ai_data = res.json()
                content = json.loads(ai_data['choices'][0]['message']['content'])
                return Tariff(
                    category=category_name,
                    commission_rate=float(content.get('commission_rate', 0.12)),
                    return_fee=float(content.get('return_rate', 0.05)),
                    source="DeepSeek AI Predict"
                )
        except Exception as e:
            logger.error(f"Ошибка обращения к DeepSeek: {e}")
        return self.tariffs['default']

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
            'Источник данных': t.source
        } for k, t in self.tariffs.items()])

# ============================================================================
# БЛОК 3: ВЕКТОРИЗОВАННЫЙ ФИНАНСОВЫЙ ДВИЖОК C СТРАТЕГИЕЙ ЦЕН
# ============================================================================
class VectorizedEnginePRO:
    @staticmethod
    def run_calculations(df: pd.DataFrame, tax_system: TaxSystem, manager: HybridTariffManager) -> pd.DataFrame:
        if df.empty: return df
        
        # Предварительная очистка от дубликатов и кракозябр
        if 'artikul' in df.columns:
            df['artikul'] = df['artikul'].astype(str).apply(fix_double_utf8)
        if 'category' in df.columns:
            df['category'] = df['category'].astype(str).apply(fix_double_utf8)

        for c in ['selling_price', 'cogs', 'weight_kg', 'length_cm', 'width_cm', 'height_cm', 'packaging_cost', 'marketing_budget_per_unit', 'daily_sales', 'stock_depth_days']:
            if c not in df.columns: df[c] = 0.0
            df[c] = pd.to_numeric(df[c], errors='coerce').fillna(0.0)

        comm_rates, min_comms, magma_bases, magma_kgs, acq_fees, ret_fees = [], [], [], [], [], []
        for cat in df.get('category', ['default'] * len(df)):
            t = manager.get_best_tariff(cat)
            comm_rates.append(t.commission_rate)
            min_comms.append(t.min_commission)
            magma_bases.append(t.magma_base)
            magma_kgs.append(t.magma_per_kg)
            acq_fees.append(t.acquiring_fee)
            ret_fees.append(t.return_fee)

        comm_rates = np.array(comm_rates)
        acq_fees = np.array(acq_fees)

        vol_weight = (df['length_cm'] * df['width_cm'] * df['height_cm']) / 5000.0
        df['billable_weight'] = np.maximum(df['weight_kg'], vol_weight)
        df['billable_weight'] = np.ceil(df['billable_weight'] * 2) / 2

        df['commission'] = np.maximum(df['selling_price'] * comm_rates, min_comms)
        df['last_mile_cost'] = np.clip(df['selling_price'] * 0.045, 60.0, 400.0)
        df['first_mile_cost'] = magma_bases + (df['billable_weight'] * magma_kgs)
        
        df['acquiring_cost'] = df['selling_price'] * acq_fees
        df['return_cost'] = (150.0 + (df['selling_price'] * 0.30)) * ret_fees
        df['pick_pack_cost'] = 35.0
        df['warehouse_cost'] = (df['stock_depth_days'] * df['daily_sales']) * 0.50

        df['fixed_operational_costs'] = (
            df['cogs'] + df['first_mile_cost'] + df['pick_pack_cost'] + 
            df['packaging_cost'] + df['return_cost'] + df['marketing_budget_per_unit'] + df['warehouse_cost']
        )

        df['pre_tax_expenses'] = df['fixed_operational_costs'] + df['commission'] + df['last_mile_cost'] + df['acquiring_cost']

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

        df['total_expenses'] = df['pre_tax_expenses'] + df['tax_cost']
        df['gross_profit'] = df['selling_price'] - df['total_expenses']
        df['margin_percent'] = np.where(df['selling_price'] > 0, (df['gross_profit'] / df['selling_price']) * 100, 0.0)

        # РЕКОМЕНДОВАННЫЕ ЦЕНЫ С ЦЕЛЕВОЙ ДОХОДНОСТЬЮ
        tax_factor = tax_system.rate if tax_system.base == "revenue" else 0.0
        variable_fees_share = comm_rates + 0.045 + acq_fees + tax_factor
        denom = 1.0 - variable_fees_share
        denom = np.where(denom <= 0.01, 0.5, denom)

        df['rec_price_min'] = df['fixed_operational_costs'] / denom
        df['rec_price_15'] = df['fixed_operational_costs'] / (denom - 0.15)
        df['rec_price_25'] = df['fixed_operational_costs'] / (denom - 0.25)

        money_columns = ['commission', 'last_mile_cost', 'first_mile_cost', 'acquiring_cost', 
                         'return_cost', 'pre_tax_expenses', 'tax_cost', 'total_expenses', 
                         'gross_profit', 'rec_price_min', 'rec_price_15', 'rec_price_25']
        for col in money_columns:
            df[col] = df[col].apply(money_round)

        return df

# ============================================================================
# БЛОК 4: ЭКСПОРТ В EXCEL С ЖИВЫМИ ФОРМУЛАМИ ПЕРЕСЧЕТА
# ============================================================================
class ExcelDynamicExporter:
    @staticmethod
    def export(df: pd.DataFrame) -> bytes:
        if not OPENPYXL_AVAILABLE: return b""
        wb = Workbook()
        
        # Лист 1: Сводная панель KPI
        ws_dash = wb.active
        ws_dash.title = "📊 Дашборд"
        ws_dash.cell(1, 1, "Сводный финансовый аналитический отчет").font = Font(size=14, bold=True)
        ws_dash.cell(3, 1, "Всего SKU в обработке:")
        ws_dash.cell(3, 2, len(df))
        ws_dash.cell(4, 1, "Суммарная плановая выручка:")
        ws_dash.cell(4, 2, f"=SUM('Расчет экономики'!C2:C{len(df)+1})")
        ws_dash.cell(5, 1, "Суммарная плановая прибыль:")
        ws_dash.cell(5, 2, f"=SUM('Расчет экономики'!L2:L{len(df)+1})")
        ws_dash.column_dimensions['A'].width = 30

        # Лист 2: Полноформатная интерактивная модель
        ws = wb.create_sheet("Расчет экономики")
        headers = ['Артикул', 'Категория', 'Цена продажи (Редактируемая)', 'Себестоимость (Закупка)', 
                   'Комиссия маркетплейса', 'Магистраль (Логистика)', 'Последняя миля', 
                   'Банковский эквайринг', 'Процент возвратов/Брак', 'Расчетный налог', 
                   'Итого расходов (Формула)', 'Чистая прибыль (Формула)', 'Текущая маржа, % (Формула)', 
                   'МИН. ЦЕНА (Порог маржи 0%)', 'ОПТИМАЛЬНАЯ ЦЕНА (Цель 15% маржи)', 'МАКСИМАЛЬНАЯ ЦЕНА (Цель 25% маржи)']
        
        for col_idx, text in enumerate(headers, 1):
            cell = ws.cell(1, col_idx, text)
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill(start_color="1F4E78", fill_type="solid")
            cell.alignment = Alignment(horizontal="center", wrap_text=True)

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
            
            # Интерактивные формулы Excel!
            ws.cell(r, 11, f"=SUM(D{r}:J{r})").border = data_border
            ws.cell(r, 12, f"=C{r}-K{r}").border = data_border
            ws.cell(r, 13, f"=IF(C{r}>0, (L{r}/C{r})*100, 0)").border = data_border
            
            # Границы цен для менеджеров
            ws.cell(r, 14, float(row['rec_price_min'])).border = data_border
            ws.cell(r, 15, float(row['rec_price_15'])).border = data_border
            ws.cell(r, 16, float(row['rec_price_25'])).border = data_border

            # Цветовое кодирование ценовых блоков стратегий
            ws.cell(r, 14).fill = PatternFill(start_color="F2DCDB", fill_type="solid")
            ws.cell(r, 15).fill = PatternFill(start_color="E2EFDA", fill_type="solid")
            ws.cell(r, 16).fill = PatternFill(start_color="D9E1F2", fill_type="solid")

        # Условное форматирование прибыли (Зеленый / Красный в зависимости от знака)
        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        ws.conditional_formatting.add(f"L2:L{len(df)+1}", CellIsRule(operator='greaterThan', formula=['0'], fill=green_fill))
        ws.conditional_formatting.add(f"L2:L{len(df)+1}", CellIsRule(operator='lessThan', formula=['0'], fill=red_fill))

        for col in range(1, len(headers) + 1):
            ws.column_dimensions[get_column_letter(col)].width = 22

        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.getvalue()

# ============================================================================
# БЛОК 5: МОДУЛЬ АВТОМАТИЧЕСКОЙ СИНХРОНИЗАЦИИ ЦЕН ПО API С КАБИНЕТОМ ZAPSTORE
# ============================================================================
class YandexMarketApiSync:
    """Выгрузка скорректированных цен в кабинет Яндекс Маркета по API."""
    @staticmethod
    def update_prices(business_id: str, api_key: str, price_data: list) -> Tuple[bool, str]:
        url = f"[link removed]{business_id}/offers/update-prices"
        headers = {
            "Authorization": f"OAuth {api_key}",
            "Content-Type": "application/json"
        }
        
        offers_payload = []
        for item in price_data:
            offers_payload.append({
                "offerId": str(item['artikul']),
                "price": {
                    "value": float(item['new_price']),
                    "currencyId": "RUR"
                }
            })
            
        payload = {"offers": offers_payload}
        try:
            response = requests.post(url, json=payload, headers=headers, timeout=10)
            if response.status_code == 200:
                res_json = response.json()
                if res_json.get("status") == "ERROR":
                    errors = res_json.get("errors", [{"message": "Неизвестная ошибка API Маркета"}])
                    return False, errors[0].get("message")
                return True, "Цены успешно обновлены и отправлены на модерацию в ЛК ZapStore."
            else:
                return False, f"Ошибка API ({response.status_code}): {response.text}"
        except Exception as e:
            return False, f"Сбой сетевого подключения к шлюзу партнерского API: {e}"

# ============================================================================
# БЛОК 6: STREAMLIT ИНТЕРФЕЙС УПРАВЛЕНИЯ И ЗАПУСКА
# ============================================================================
def main():
    st.set_page_config(page_title=APP_NAME, page_icon="📈", layout="wide")
    st.title(f"📈 {APP_NAME} v{APP_VERSION} — Корпоративный симулятор цен")
    st.caption("Автоматизация маркетплейсов: Векторизованные финансовые симуляции, интерактивный расчёт рекомендованных цен и синхронизация по API.")

    # Корпоративные константы из файлов конфигурации ZapStore
    COMPANY_API_KEY = "ACMA:baYKVsVh7vORZYIZLLvZviviZAxfjcRmdrariFBH:e755690c"
    COMPANY_CAMPAIGN_ID = "78311459"
    COMPANY_BUSINESS_ID = "93193868"

    if 'tm' not in st.session_state: st.session_state.tm = HybridTariffManager()
    if 'main_df' not in st.session_state:
        st.session_state.main_df = pd.DataFrame([{
            'artikul': 'PART-7831', 'category': 'автозапчасти', 'selling_price': 3990.0, 'cogs': 1800.0,
            'weight_kg': 1.5, 'length_cm': 25.0, 'width_cm': 15.0, 'height_cm': 10.0,
            'packaging_cost': 40.0, 'marketing_budget_per_unit': 200.0, 'daily_sales': 4, 'stock_depth_days': 30
        }])

    # Раздел гибридного управления источниками тарификации
    with st.expander("🌐 Настройка гибридного получения тарифов (API Яндекс Маркет / ИИ DeepSeek)", expanded=False):
        c1, c2, c3 = st.columns(3)
        with c1:
            st.markdown("**🔌 Подключение к Партнерскому API**")
            oauth = st.text_input("OAuth Токен доступа", type="password", value=COMPANY_API_KEY)
            camp_id = st.text_input("Идентификатор кампании (campaignId)", value=COMPANY_CAMPAIGN_ID)
            if st.button("Синхронизировать сетку тарифов с API"):
                if st.session_state.tm.fetch_from_yandex_market_api(oauth, camp_id):
                    st.success("Успешное обновление тарифов из ЛК Яндекс Маркета!")
                else:
                    st.error("Ошибка запроса к партнерскому шлюзу API.")
        with c2:
            st.markdown("**🧠 Прогнозирование через DeepSeek AI**")
            ds_key = st.text_input("DeepSeek API Key", type="password")
            ai_cat = st.text_input("Указать категорию для AI анализа", value="Тормозные колодки")
            if st.button("Сгенерировать тариф ИИ"):
                if ds_key:
                    predicted = st.session_state.tm.fetch_from_deepseek_ai(ds_key, ai_cat)
                    st.session_state.tm.tariffs[ai_cat.lower()] = predicted
                    st.success(f"ИИ предсказал базовую комиссию: {predicted.commission_rate*100}%")
                else:
                    st.warning("Необходимо указать API ключ авторизации DeepSeek.")
        with c3:
            st.markdown("**📊 Текущая тарифная сетка**")
            st.dataframe(st.session_state.tm.to_dataframe(), use_container_width=True)

    tax_label = st.selectbox("Налоговый режим бизнеса:", ["УСН 6% (доходы)", "УСН 15% (доходы-расходы)", "ОСН (общая с НДС 20%)"])
    current_tax = TaxSystem.by_label(tax_label)

    st.markdown("### 📝 Интерактивный симулятор товарной матрицы")
    st.info("💡 Кликните дважды на ячейку **selling_price** (Текущая цена) или **cogs** (Закупочная стоимость) для изменения параметров — вся модель пересчитается мгновенно.")

    # Интерактивный редактор товарной матрицы
    edited_df = st.data_editor(st.session_state.main_df, num_rows="dynamic", use_container_width=True)

    if not edited_df.empty:
        # Запуск векторизованного пересчета измененных параметров
        calculated_df = VectorizedEnginePRO.run_calculations(edited_df, current_tax, st.session_state.tm)
        st.session_state.main_df = edited_df

        st.markdown("### 🗠 Стратегический анализ цен, маржи и ценовых коридоров")
        st.dataframe(calculated_df[[
            'artikul', 'category', 'selling_price', 'cogs', 'pre_tax_expenses', 
            'gross_profit', 'margin_percent', 'rec_price_min', 'rec_price_15', 'rec_price_25'
        ]], use_container_width=True)

        # Вывод ключевых показателей
        c_min_p = calculated_df['rec_price_min'].mean()
        c_15_p = calculated_df['rec_price_15'].mean()
        c_25_p = calculated_df['rec_price_25'].mean()
        
        met1, met2, met3 = st.columns(3)
        met1.metric("Порог безубыточности (средний)", f"{c_min_p:,.2f} ₽", "Маржа 0%")
        met2.metric("Рекомендованная цена (Оптимум)", f"{c_15_p:,.2f} ₽", "Чистая маржа 15%")
        met3.metric("Цена максимальной доходности", f"{c_25_p:,.2f} ₽", "Чистая маржа 25%")

        # Секция выгрузки в Excel
        st.markdown("---")
        st.subheader("📥 Выгрузка интерактивной Excel модели")
        if OPENPYXL_AVAILABLE:
            excel_bytes = ExcelDynamicExporter.export(calculated_df)
            st.download_button(
                label="⬇️ СКАЧАТЬ ВЫСОКОИНФОРМАТИВНЫЙ EXCEL ОТЧЕТ С ЖИВЫМИ ФОРМУЛАМИ (.XLSX)",
                data=excel_bytes,
                file_name=f"Market_Price_Strategy_{datetime.now().strftime('%d_%m_%Y')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        else:
            st.warning("Установите библиотеку openpyxl.")

        # Блок автоматической отправки цен по API в ЛК ZapStore
        st.markdown("---")
        st.subheader("📡 Модуль автоматической синхронизации цен с Яндекс Маркетом")
        selected_strategy = st.selectbox("Выберите ценовую стратегию для пакетной отправки в магазин ZapStore по API:", 
                                         ["Текущая установленная цена", "Минимальная цена (Безубыточность)", "Оптимальная цена (15% маржа)", "Максимальная цена (25% маржа)"])
        
        if st.button("🚀 ОТПРАВИТЬ ОБНОВЛЕННЫЕ ЦЕНЫ НА ЯНДЕКС МАРКЕТ", type="primary", use_container_width=True):
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
            
            with st.spinner("Передача пакета данных по защищенному API шлюзу..."):
                success, msg = YandexMarketApiSync.update_prices(
                    business_id=COMPANY_BUSINESS_ID,
                    api_key=COMPANY_API_KEY,
                    price_data=price_data_to_send
                )
                if success:
                    st.success(f"Выгрузка успешно завершена! {msg}")
                else:
                    st.error(f"Ошибка транзакции цен: {msg}")

if __name__ == "__main__":
    main()
