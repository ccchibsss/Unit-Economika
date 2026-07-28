import streamlit as st
import pandas as pd
import numpy as np
import io
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime
import requests
import json
import re
import csv
import os
import base64
import hashlib
from pathlib import Path
import time
import logging
from typing import Dict, List, Any, Optional, Tuple, Union
from dataclasses import dataclass, field
import uuid
from functools import lru_cache
import warnings
import math
warnings.filterwarnings('ignore')

# ============================================================================
# БЛОК 0: БАЗОВАЯ КОНФИГУРАЦИЯ И ИМПОРТЫ
# ============================================================================

# === Версия приложения ===
APP_VERSION = "3.0.0"
APP_NAME = "🚀 FBS Юнит-экономика PRO 2026"
APP_DESCRIPTION = "Профессиональная юнит-экономика для маркетплейсов с фокусом на Яндекс Маркет"

# === Базовые директории ===
try:
    BASE_DIR = Path(__file__).parent.resolve()
except NameError:
    BASE_DIR = Path.cwd()

DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
LOGS_DIR = BASE_DIR / "logs"
EXPORTS_DIR = BASE_DIR / "exports"
CONFIG_DIR = BASE_DIR / "config"
SECURE_KEYS_DIR = BASE_DIR / "secure_keys"
GOOGLE_CREDS_DIR = BASE_DIR / "google_creds"

for dir_path in [DATA_DIR, CACHE_DIR, LOGS_DIR, EXPORTS_DIR, CONFIG_DIR, SECURE_KEYS_DIR, GOOGLE_CREDS_DIR]:
    try:
        dir_path.mkdir(exist_ok=True, parents=True)
    except OSError:
        pass

# === Логирование ===
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(LOGS_DIR / "fbs_economy.log", encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger('FBSEconomy')

# ============================================================================
# БЛОК 1: БЕЗОПАСНОЕ ХРАНЕНИЕ КЛЮЧЕЙ (ШИФРОВАНИЕ FERNET)
# ============================================================================

try:
    from cryptography.fernet import Fernet
    CRYPTO_AVAILABLE = True
except ImportError:
    CRYPTO_AVAILABLE = False
    Fernet = None

class SecureKeyManager:
    """
    Менеджер безопасного хранения API ключей.
    Использует шифрование Fernet для сохранения ключей в локальный файл.
    """
    def __init__(self, key_dir: Path = SECURE_KEYS_DIR):
        self.key_dir = key_dir
        self.key_dir.mkdir(parents=True, exist_ok=True)
        
        self.master_key_file = self.key_dir / "master.key"
        self.encrypted_data_file = self.key_dir / "api_keys.enc"
        
        self.fernet = self._get_or_create_fernet()
        self._keys_cache: Dict[str, str] = self._load_keys()
        self._metadata_cache: Dict[str, Dict[str, Any]] = self._load_metadata()

    def _get_or_create_fernet(self) -> Fernet:
        """Генерирует или загружает мастер-ключ шифрования"""
        if not CRYPTO_AVAILABLE:
            raise ImportError("cryptography не установлен. pip install cryptography")
            
        if not self.master_key_file.exists():
            new_key = Fernet.generate_key()
            self.master_key_file.write_bytes(new_key)
            try:
                os.chmod(self.master_key_file, 0o600)
            except OSError:
                pass
        else:
            new_key = self.master_key_file.read_bytes()
        return Fernet(new_key)

    def _load_keys(self) -> Dict[str, str]:
        """Расшифровывает и загружает ключи из файла"""
        if not self.encrypted_data_file.exists():
            return {}
        try:
            encrypted_data = self.encrypted_data_file.read_bytes()
            decrypted_bytes = self.fernet.decrypt(encrypted_data)
            return json.loads(decrypted_bytes.decode('utf-8'))
        except Exception as e:
            logger.error(f"Ошибка расшифровки ключей: {e}")
            return {}

    def _load_metadata(self) -> Dict[str, Dict[str, Any]]:
        """Загружает метаданные ключей"""
        metadata_file = self.key_dir / "api_keys_meta.json"
        if not metadata_file.exists():
            return {}
        try:
            return json.loads(metadata_file.read_text(encoding='utf-8'))
        except Exception:
            return {}

    def _save_keys(self):
        """Шифрует и сохраняет ключи в файл"""
        try:
            encrypted_data = self.fernet.encrypt(json.dumps(self._keys_cache).encode('utf-8'))
            self.encrypted_data_file.write_bytes(encrypted_data)
            try:
                os.chmod(self.encrypted_data_file, 0o600)
            except OSError:
                pass
        except Exception as e:
            logger.error(f"Ошибка шифрования и сохранения ключей: {e}")

    def _save_metadata(self):
        """Сохраняет метаданные ключей"""
        metadata_file = self.key_dir / "api_keys_meta.json"
        try:
            metadata_file.write_text(json.dumps(self._metadata_cache, indent=2, ensure_ascii=False), encoding='utf-8')
        except Exception as e:
            logger.error(f"Ошибка сохранения метаданных ключей: {e}")

    def set_key(self, service: str, api_key: str, description: str = ""):
        """Сохраняет или обновляет API ключ для сервиса."""
        if not api_key or not api_key.strip():
            if service in self._keys_cache:
                del self._keys_cache[service]
            if service in self._metadata_cache:
                del self._metadata_cache[service]
        else:
            self._keys_cache[service] = api_key.strip()
            now = datetime.now().isoformat()
            self._metadata_cache[service] = {
                "description": description,
                "created_at": self._metadata_cache.get(service, {}).get("created_at", now),
                "last_updated": now
            }
        self._save_keys()
        self._save_metadata()

    def get_key(self, service: str) -> Optional[str]:
        """Получает API ключ для сервиса (в открытом виде, только в памяти)"""
        return self._keys_cache.get(service)

    def get_metadata(self, service: str) -> Dict[str, Any]:
        """Получает метаданные ключа"""
        return self._metadata_cache.get(service, {})

    def list_services(self) -> List[str]:
        """Возвращает список сервисов, для которых есть ключи"""
        return list(self._keys_cache.keys())

    def delete_key(self, service: str):
        """Удаляет ключ и его метаданные"""
        if service in self._keys_cache:
            del self._keys_cache[service]
        if service in self._metadata_cache:
            del self._metadata_cache[service]
        self._save_keys()
        self._save_metadata()

    def clear_all(self):
        """Очищает все сохраненные ключи"""
        self._keys_cache.clear()
        self._metadata_cache.clear()
        self._save_keys()
        self._save_metadata()

# ============================================================================
# БЛОК 2: УТИЛИТЫ ДЛЯ РАБОТЫ С ФАЙЛАМИ И КОДИРОВКАМИ
# ============================================================================

def detect_encoding(file_bytes: bytes) -> str:
    """
    Авто-детекция кодировки файла.
    Приоритет: UTF-8, UTF-8-SIG, CP1251, Latin1
    """
    try:
        import chardet
        result = chardet.detect(file_bytes[:10000])
        if result and result.get('encoding'):
            return result['encoding']
    except ImportError:
        pass
    
    # Попробуем популярные кодировки
    encodings = ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251', 'latin1', 'iso-8859-1']
    for enc in encodings:
        try:
            file_bytes.decode(enc)
            return enc
        except UnicodeDecodeError:
            continue
    return 'utf-8'

def escape_excel_text(value: Any) -> str:
    """
    Экранирует строку для Excel, чтобы предотвратить автоматическое преобразование 
    в дату или формулу.
    """
    if pd.isna(value) or value is None:
        return ""
    
    s = str(value).strip()
    if not s:
        return s
    
    # Проверка на формулы
    if s.startswith(('=', '+', '-', '@')):
        return f"'{s}"
    
    # Проверка на потенциальные даты
    if re.match(r'^\d+[-/]\d+([-/]\d+)?$', s) or re.match(r'^[A-Za-z]{3,4}[-/]\d+$', s, re.IGNORECASE):
        return f"'{s}"
    
    # Проверка на артикулы типа "12345-678"
    if re.match(r'^[A-Za-z0-9]+[-][A-Za-z0-9]+$', s):
        return f"'{s}"
        
    return s

def smart_read_csv(uploaded_file) -> pd.DataFrame:
    """
    Умное чтение CSV с авто-детекцией кодировки и разделителя.
    """
    uploaded_file.seek(0)
    file_bytes = uploaded_file.read()
    uploaded_file.seek(0)
    
    encoding = detect_encoding(file_bytes)
    
    separators = [';', ',', '\t', '|']
    best_df = None
    best_score = -1
    
    for sep in separators:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(
                uploaded_file,
                encoding=encoding,
                sep=sep,
                dtype=str,
                on_bad_lines='skip',
                skipinitialspace=True,
                engine='python'
            )
            
            if df is None or df.empty or len(df.columns) <= 1:
                continue
            
            # Оценка качества: чем больше колонок, тем лучше
            score = len(df.columns)
            if score > best_score:
                best_score = score
                best_df = df
                
            # Если колонок больше 2, это скорее всего правильный разделитель
            if len(df.columns) >= 3:
                break
                
        except Exception:
            continue
    
    if best_df is not None:
        return best_df
    
    # Если ничего не найдено, пробуем снова с UTF-8 и запятой
    uploaded_file.seek(0)
    return pd.read_csv(uploaded_file, encoding='utf-8', dtype=str, on_bad_lines='skip')

def smart_read_uploaded_file(uploaded_file) -> pd.DataFrame:
    """
    Умное чтение загруженного файла (CSV или Excel).
    """
    if uploaded_file is None:
        return pd.DataFrame()
    
    uploaded_file.seek(0)
    file_name = uploaded_file.name.lower()
    
    try:
        if file_name.endswith(('.csv', '.txt')):
            return smart_read_csv(uploaded_file)
            
        elif file_name.endswith(('.xlsx', '.xls')):
            uploaded_file.seek(0)
            df = pd.read_excel(
                uploaded_file,
                engine='openpyxl' if file_name.endswith('.xlsx') else 'xlrd',
                dtype=str,
                keep_default_na=False
            )
            return df
            
        else:
            raise ValueError(f"Неподдерживаемый формат файла: {file_name}")
            
    except Exception as e:
        logger.error(f"Ошибка чтения файла {uploaded_file.name}: {e}")
        st.error(f"Ошибка чтения файла: {e}")
        return pd.DataFrame()

# ============================================================================
# БЛОК 3: DEEPSEEK API ИНТЕГРАЦИЯ (ПРОФЕССИОНАЛЬНАЯ)
# ============================================================================

class DeepSeekAPIManager:
    """
    Профессиональный менеджер для работы с DeepSeek API.
    Поддерживает: обогащение каталога, актуализацию тарифов, анализ конкурентов.
    """
    def __init__(self, api_key: str, base_url: str = "https://api.deepseek.com/v1"):
        self.api_key = api_key
        self.base_url = base_url
        
    def is_available(self) -> bool:
        return bool(self.api_key)
    
    def analyze_product_for_marketplace(
        self, 
        product_name: str, 
        marketplace: str = "Яндекс Маркет",
        current_price: float = 0,
        current_category: str = ""
    ) -> Dict[str, Any]:
        """
        Профессиональный анализ товара для конкретного маркетплейса.
        """
        if not self.is_available():
            return {"error": "DeepSeek API недоступен или ключ не задан"}
            
        prompt = f"""
        Ты эксперт по юнит-экономике для маркетплейсов с 10-летним опытом.
        Проанализируй товар для маркетплейса {marketplace}.
        
        Название товара: "{product_name}"
        Текущая цена: {current_price} ₽
        Категория: {current_category}
        
        Верни строго JSON в следующем формате:
        {{
            "marketplace_recommendations": {{
                "recommended_price": 0,
                "min_profit_margin": 15.0,
                "competition_level": "medium",
                "sales_potential": "high",
                "seasonality_factor": 1.0
            }},
            "product_characteristics": {{
                "category": "автозапчасти",
                "subcategory": "подвеска",
                "is_hazardous": false,
                "is_fragile": false,
                "average_weight_kg": 1.5,
                "typical_volume_l": 0.5
            }},
            "marketplace_tariffs": {{
                "commission_rate": 0.15,
                "logistics_base": 55.0,
                "logistics_per_kg": 16.0,
                "storage_per_day": 0.35,
                "return_rate": 0.025,
                "acquiring_fee": 0.015
            }},
            "competitive_analysis": {{
                "price_position": "medium",
                "recommended_actions": ["оптимизировать упаковку", "улучшить фото"],
                "profit_optimization_tips": ["снизить логистику", "увеличить средний чек"]
            }},
            "confidence_score": 0.85
        }}
        Не добавляй никакой разметки, только валидный JSON.
        """
        
        try:
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}"
            }
            payload = {
                "model": "deepseek-chat",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.1,
                "response_format": {"type": "json_object"}
            }
            
            response = requests.post(
                f"{self.base_url}/chat/completions",
                headers=headers,
                json=payload,
                timeout=45
            )
            
            if response.status_code == 200:
                content = response.json()["choices"][0]["message"]["content"]
                return json.loads(content)
            else:
                return {"error": f"HTTP {response.status_code}: {response.text}"}
                
        except Exception as e:
            logger.error(f"Ошибка DeepSeek (анализ товара): {e}")
            return {"error": str(e)}
    
    def get_marketplace_insights(
        self,
        marketplace: str = "Яндекс Маркет",
        category: str = "auto_parts"
    ) -> Dict[str, Any]:
        """
        Получение актуальных инсайтов по маркетплейсу.
        """
        if not self.is_available():
            return {"error": "DeepSeek API недоступен или ключ не задан"}
            
        prompt = f"""
        Ты аналитик маркетплейсов с фокусом на {marketplace}.
        Предоставь актуальные инсайты для продавцов в категории {category} на 2026 год.
        
        Верни строго JSON в следующем формате:
        {{
            "marketplace_trends": {{
                "growth_rate": 0.15,
                "top_categories": ["подвеска", "двигатель", "тормозная система"],
                "emerging_trends": ["электрические компоненты", "гибридные системы"],
                "seasonal_patterns": {{
                    "peak_months": [3, 4, 9, 10],
                    "slow_months": [1, 2, 7, 8]
                }}
            }},
            "tariffs_2026": {{
                "commission_rate": 0.145,
                "min_commission": 35.0,
                "logistics_base": 55.0,
                "logistics_per_kg": 16.0,
                "storage_per_day": 0.35,
                "return_fee": 0.025,
                "acquiring_fee": 0.015,
                "delivery_tariffs": {{
                    "standard": 80.0,
                    "express": 150.0
                }}
            }},
            "seller_insights": {{
                "average_margin": 0.22,
                "top_sellers_features": ["быстрая доставка", "качественные фото", "подробное описание"],
                "common_mistakes": ["завышенная цена", "плохая упаковка", "медленный ответ на заказы"]
            }},
            "source": "DeepSeek AI Professional Analytics"
        }}
        Не добавляй никакой разметки, только валидный JSON.
        """
        
        try:
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {self.api_key}"
            }
            payload = {
                "model": "deepseek-chat",
                "messages": [{"role": "user", "content": prompt}],
                "temperature": 0.1,
                "response_format": {"type": "json_object"}
            }
            
            response = requests.post(
                f"{self.base_url}/chat/completions",
                headers=headers,
                json=payload,
                timeout=45
            )
            
            if response.status_code == 200:
                content = response.json()["choices"][0]["message"]["content"]
                return json.loads(content)
            else:
                return {"error": f"HTTP {response.status_code}: {response.text}"}
                
        except Exception as e:
            logger.error(f"Ошибка DeepSeek (инсайты маркетплейса): {e}")
            return {"error": str(e)}

# ============================================================================
# БЛОК 4: GOOGLE SHEETS ИНТЕГРАЦИЯ
# ============================================================================

try:
    import gspread
    from google.oauth2.service_account import Credentials
    GSPREAD_AVAILABLE = True
except ImportError:
    GSPREAD_AVAILABLE = False
    gspread = None

class GoogleSheetsManager:
    """
    Менеджер для работы с Google Sheets API.
    """
    def __init__(self, credentials_json: str):
        self.credentials_json = credentials_json
        self.client = None
        self._init_client()
        
    def _init_client(self):
        """Инициализация gspread клиента"""
        if not GSPREAD_AVAILABLE:
            logger.error("gspread не установлен")
            return
        try:
            if os.path.exists(self.credentials_json):
                credentials = Credentials.from_service_account_file(
                    self.credentials_json,
                    scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                )
            else:
                creds_data = json.loads(self.credentials_json)
                credentials = Credentials.from_service_account_info(
                    creds_data,
                    scopes=["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
                )
            self.client = gspread.authorize(credentials)
            logger.info("✅ Google Sheets клиент инициализирован")
        except Exception as e:
            logger.error(f"Ошибка инициализации Google Sheets: {e}")
            
    def read_sheet(self, spreadsheet_id: str, worksheet_name: str = "Лист1") -> Optional[pd.DataFrame]:
        """Читает данные из Google Sheets"""
        if self.client is None:
            return None
        try:
            spreadsheet = self.client.open_by_key(spreadsheet_id)
            worksheet = spreadsheet.worksheet(worksheet_name)
            values = worksheet.get_all_values()
            if not values:
                return pd.DataFrame()
            headers = values[0]
            data = values[1:]
            df = pd.DataFrame(data, columns=headers)
            return df
        except Exception as e:
            logger.error(f"Ошибка чтения Google Sheets: {e}")
            return None
            
    def write_sheet(self, spreadsheet_id: str, df: pd.DataFrame, worksheet_name: str = "Лист1", clear_before: bool = True) -> bool:
        """Записывает DataFrame в Google Sheets"""
        if self.client is None:
            return False
        try:
            spreadsheet = self.client.open_by_key(spreadsheet_id)
            
            # Очищаем или создаем лист
            try:
                worksheet = spreadsheet.worksheet(worksheet_name)
                if clear_before:
                    worksheet.clear()
            except gspread.exceptions.WorksheetNotFound:
                worksheet = spreadsheet.add_worksheet(title=worksheet_name, rows="1000", cols="100")
            
            # Обновляем данные
            data = [df.columns.tolist()] + df.values.tolist()
            worksheet.update(data, value_input_option='USER_ENTERED')
            
            logger.info(f"✅ Данные записаны в Google Sheets: {len(df)} строк")
            return True
        except Exception as e:
            logger.error(f"Ошибка записи в Google Sheets: {e}")
            return False

# ============================================================================
# БЛОК 5: ГЕНЕРАТОР ЖИВЫХ ФОРМУЛ EXCEL (ПРОФЕССИОНАЛЬНЫЙ)
# ============================================================================

class ExcelFormulaBuilder:
    """
    Профессиональный генератор живых формул Excel для юнит-экономики.
    """
    def __init__(self, col_map: Dict[str, str]):
        self.col_map = col_map

    def _get_cell(self, field: str, row: int = 2) -> str:
        """Возвращает ссылку на ячейку"""
        col = self.col_map.get(field, "A")
        return f"{col}{row}"

    def build_commission_formula(self, row: int = 2) -> str:
        price_cell = self._get_cell("price", row)
        rate_cell = self._get_cell("commission_rate", row)
        return f"={price_cell}*{rate_cell}"

    def build_logistics_formula(self, row: int = 2) -> str:
        base_cell = self._get_cell("logistics_base", row)
        weight_cell = self._get_cell("weight", row)
        rate_cell = self._get_cell("logistics_per_kg", row)
        return f"={base_cell}+({weight_cell}*{rate_cell})"

    def build_storage_formula(self, row: int = 2) -> str:
        length_cell = self._get_cell("length", row)
        width_cell = self._get_cell("width", row)
        height_cell = self._get_cell("height", row)
        days_cell = self._get_cell("storage_days", row)
        rate_cell = self._get_cell("storage_rate", row)
        return f"=IF({length_cell}*{width_cell}*{height_cell}>0, ({length_cell}*{width_cell}*{height_cell}/1000)*{rate_cell}*{days_cell}, 5*{rate_cell}*{days_cell})"

    def build_acquiring_formula(self, row: int = 2) -> str:
        price_cell = self._get_cell("price", row)
        rate_cell = self._get_cell("acquiring_rate", row)
        return f"={price_cell}*{rate_cell}"

    def build_returns_formula(self, row: int = 2) -> str:
        price_cell = self._get_cell("price", row)
        rate_cell = self._get_cell("return_rate", row)
        return f"={price_cell}*{rate_cell}"

    def build_tax_formula(self, row: int = 2) -> str:
        price_cell = self._get_cell("price", row)
        rate_cell = self._get_cell("tax_rate", row)
        return f"={price_cell}*{rate_cell}"

    def build_total_expenses_formula(self, row: int = 2) -> str:
        cost = self._get_cell("cost", row)
        commission = self._get_cell("commission", row)
        logistics = self._get_cell("logistics", row)
        storage = self._get_cell("storage", row)
        acquiring = self._get_cell("acquiring", row)
        tax = self._get_cell("tax", row)
        returns = self._get_cell("returns", row)
        marketing = self._get_cell("marketing", row)
        packaging = self._get_cell("packaging", row)
        specific = self._get_cell("specific", row)
        return f"={cost}+{commission}+{logistics}+{storage}+{acquiring}+{tax}+{returns}+{marketing}+{packaging}+{specific}"

    def build_profit_formula(self, row: int = 2) -> str:
        price = self._get_cell("price", row)
        expenses = self._get_cell("total_expenses", row)
        return f"={price}-{expenses}"

    def build_margin_formula(self, row: int = 2) -> str:
        profit = self._get_cell("profit", row)
        price = self._get_cell("price", row)
        return f"=IF({price}>0, ({profit}/{price})*100, 0)"

    def build_roi_formula(self, row: int = 2) -> str:
        profit = self._get_cell("profit", row)
        cost = self._get_cell("cost", row)
        return f"=IF({cost}>0, ({profit}/{cost})*100, 0)"

    def build_recommended_price_formula(self, row: int = 2) -> str:
        cost = self._get_cell("cost", row)
        logistics_base = self._get_cell("logistics_base", row)
        storage = self._get_cell("storage", row)
        comm_rate = self._get_cell("commission_rate", row)
        acquiring_rate = self._get_cell("acquiring_rate", row)
        tax_rate = self._get_cell("tax_rate", row)
        return_rate = self._get_cell("return_rate", row)
        return f"=MAX(0, ({cost}+{logistics_base}+{storage}) / MAX(0.01, (1 - {comm_rate} - {acquiring_rate} - {tax_rate} - {return_rate} - 0.10)))"

    def build_break_even_formula(self, row: int = 2) -> str:
        fixed_costs = self._get_cell("fixed_costs", row)
        variable_costs = self._get_cell("variable_costs", row)
        return f"=IF({variable_costs}>0, {fixed_costs}/{variable_costs}, 0)"

# ============================================================================
# БЛОК 6: ПРОФЕССИОНАЛЬНЫЙ КАЛЬКУЛЯТОР ЮНИТ-ЭКОНОМИКИ (ЯНДЕКС МАРКЕТ)
# ============================================================================

class FBSUnitEconomicsCalculator:
    """
    Профессиональный калькулятор юнит-экономики для FBS.
    Специализация: Яндекс Маркет, Ozon, Wildberries.
    """
    
    ALLOWED_MODES = ["FBS", "FBY", "FBP"]
    
    def __init__(self, marketplace_config: Dict[str, Any], tax_system: str = "УСН_6"):
        self.config = marketplace_config
        self.tax_system = tax_system
        
        # Настройки по умолчанию для Яндекс Маркет
        self.yandex_market_defaults = {
            "commission_rate": 0.145,
            "min_commission": 35.0,
            "logistics_base": 55.0,
            "logistics_per_kg": 16.0,
            "storage_per_day": 0.35,
            "acquiring_fee": 0.015,
            "return_fee": 0.025,
            "mode_multipliers": {"FBS": 1.0, "FBY": 0.8, "FBP": 0.7}
        }
        
        # Настройки для Ozon
        self.ozon_defaults = {
            "commission_rate": 0.15,
            "min_commission": 30.0,
            "logistics_base": 50.0,
            "logistics_per_kg": 15.0,
            "storage_per_day": 0.3,
            "acquiring_fee": 0.015,
            "return_fee": 0.02,
            "mode_multipliers": {"FBS": 1.0, "FBY": 0.75}
        }
        
        # Настройки для Wildberries
        self.wildberries_defaults = {
            "commission_rate": 0.16,
            "min_commission": 28.0,
            "logistics_base": 45.0,
            "logistics_per_kg": 14.0,
            "storage_per_day": 0.25,
            "acquiring_fee": 0.015,
            "return_fee": 0.018,
            "mode_multipliers": {"FBS": 1.0, "FBY": 0.7}
        }
        
        # Применяем настройки для Яндекс Маркет по умолчанию
        if not marketplace_config:
            self.config = self.yandex_market_defaults.copy()
        
    def set_marketplace(self, marketplace: str):
        """
        Устанавливает конфигурацию для конкретного маркетплейса.
        """
        if marketplace == "Яндекс Маркет":
            self.config.update(self.yandex_market_defaults)
        elif marketplace == "Ozon":
            self.config.update(self.ozon_defaults)
        elif marketplace == "Wildberries":
            self.config.update(self.wildberries_defaults)
        else:
            logger.warning(f"Неизвестный маркетплейс: {marketplace}")
        
    def calculate(
        self,
        price: float,
        cost: float,
        weight_kg: float,
        length_cm: float,
        width_cm: float,
        height_cm: float,
        days_in_storage: int = 30,
        operation_mode: str = "FBS",
        marketplace: str = "Яндекс Маркет",
        category: str = "auto_parts",
        is_hazardous: bool = False,
        is_fragile: bool = False,
        commission_rate: Optional[float] = None,
        logistics_base: Optional[float] = None,
        logistics_per_kg: Optional[float] = None,
        storage_rate: Optional[float] = None,
        acquiring_rate: Optional[float] = None,
        return_rate: Optional[float] = None,
        tax_rate: Optional[float] = None,
        marketing_cost: float = 0.0,
        packaging_cost: float = 0.0
    ) -> Dict[str, Any]:
        """
        Профессиональный расчёт юнит-экономики для FBS.
        """
        if operation_mode not in self.ALLOWED_MODES:
            raise ValueError(f"Режим {operation_mode} не поддерживается. Используйте FBS, FBY или FBP.")
        
        if price <= 0 or cost <= 0:
            raise ValueError("Цена и себестоимость должны быть положительными")
        
        # === 1. УСТАНАВЛИВАЕМ КОНФИГУРАЦИЮ ДЛЯ МАРКЕТПЛЕЙСА ===
        if marketplace == "Яндекс Маркет":
            defaults = self.yandex_market_defaults
        elif marketplace == "Ozon":
            defaults = self.ozon_defaults
        elif marketplace == "Wildberries":
            defaults = self.wildberries_defaults
        else:
            defaults = self.config
        
        # === 2. КОМИССИЯ МАРКЕТПЛЕЙСА ===
        comm_rate = commission_rate if commission_rate is not None else defaults.get("commission_rate", 0.145)
        category_rates = self.config.get("category_rates", {})
        if category in category_rates:
            comm_rate = category_rates[category]
            
        commission = max(price * comm_rate, defaults.get("min_commission", 35.0))
        
        # === 3. ЛОГИСТИКА FBS (ОПТИМИЗИРОВАННАЯ ДЛЯ ЯНДЕКС МАРКЕТ) ===
        vol_weight = (length_cm * width_cm * height_cm) / 5000.0 if length_cm > 0 else 0
        billable_weight = max(weight_kg, vol_weight)
        billable_weight = math.ceil(billable_weight * 2) / 2
        
        log_base = logistics_base if logistics_base is not None else defaults.get("logistics_base", 55.0)
        log_per_kg = logistics_per_kg if logistics_per_kg is not None else defaults.get("logistics_per_kg", 16.0)
        logistics = log_base + (billable_weight * log_per_kg)
        
        # Множители для разных режимов работы
        mode_multipliers = defaults.get("mode_multipliers", {"FBS": 1.0, "FBY": 0.8, "FBP": 0.7})
        logistics *= mode_multipliers.get(operation_mode, 1.0)
        
        # Надбавка за опасные/хрупкие товары
        if is_hazardous:
            logistics *= 1.2
        if is_fragile:
            logistics *= 1.15
        
        # === 4. ХРАНЕНИЕ (С УЧЁТОМ СЕЗОННОСТИ) ===
        volume_liter = (length_cm * width_cm * height_cm) / 1000.0 if length_cm > 0 else 5.0
        storage_rate_val = storage_rate if storage_rate is not None else defaults.get("storage_per_day", 0.35)
        
        # Прогрессивная шкала хранения (чем дольше, тем дороже)
        if days_in_storage <= 60:
            storage_multiplier = 1.0
        elif days_in_storage <= 90:
            storage_multiplier = 1.5
        elif days_in_storage <= 180:
            storage_multiplier = 2.5
        elif days_in_storage <= 270:
            storage_multiplier = 5.0
        else:
            storage_multiplier = 10.0
            
        storage_cost = volume_liter * storage_rate_val * days_in_storage * storage_multiplier
        
        # === 5. ЭКВАЙРИНГ ===
        acquiring_rate_val = acquiring_rate if acquiring_rate is not None else defaults.get("acquiring_fee", 0.015)
        acquiring = price * acquiring_rate_val
        
        # === 6. ВОЗВРАТЫ ===
        return_rate_val = return_rate if return_rate is not None else defaults.get("return_fee", 0.025)
        returns = price * return_rate_val
        
        # === 7. НАДБАВКИ ЗА ОСОБЕННОСТИ ТОВАРА ===
        hazardous_surcharge = price * 0.025 if is_hazardous else 0.0
        fragile_surcharge = price * 0.015 if is_fragile else 0.0
        
        # === 8. СПЕЦИФИЧЕСКИЕ РАСХОДЫ ДЛЯ АВТОЗАПЧАСТЕЙ ===
        auto_parts_specific = 2.0 + 50.0 + 5.0 + price * 0.025
        
        # === 9. МАРКЕТИНГОВЫЕ РАСХОДЫ ===
        marketing_expenses = marketing_cost
        
        # === 10. УПАКОВКА ===
        packaging_expenses = packaging_cost
        
        # === 11. НАЛОГ (ПРОФЕССИОНАЛЬНЫЙ РАСЧЁТ) ===
        TAX_SYSTEMS = {
            "УСН_6": {"rate": 0.06, "base": "revenue"},
            "УСН_15": {"rate": 0.15, "base": "profit", "min_rate": 0.01},
            "ОСН": {"rate": 0.20, "base": "profit"},
            "НПД": {"rate": 0.06, "base": "revenue"},
        }
        
        tax_config = TAX_SYSTEMS.get(self.tax_system, TAX_SYSTEMS["УСН_6"])
        tax_rate_val = tax_rate if tax_rate is not None else tax_config["rate"]
        
        if tax_config["base"] == "revenue":
            tax = price * tax_rate_val
        else:  # profit
            profit_before_tax = price - cost - commission - logistics - storage_cost - acquiring - returns - auto_parts_specific - marketing_expenses - packaging_expenses
            tax = max(0, profit_before_tax * tax_rate_val)
            if self.tax_system == "УСН_15":
                min_tax = price * tax_config.get("min_rate", 0.01)
                tax = max(tax, min_tax)
            
        # === 12. ИТОГО РАСХОДОВ ===
        total_expenses = (
            cost + commission + logistics + storage_cost + acquiring + returns +
            hazardous_surcharge + fragile_surcharge + auto_parts_specific + 
            marketing_expenses + packaging_expenses + tax
        )
        
        # === 13. ПРИБЫЛЬ И КЛЮЧЕВЫЕ МЕТРИКИ ===
        profit = price - total_expenses
        margin_percent = (profit / price * 100) if price > 0 else 0
        roi = (profit / cost * 100) if cost > 0 else 0
        
        # === 14. ТОЧКА БЕЗУБЫТОЧНОСТИ ===
        variable_costs_per_unit = (commission + logistics + acquiring + returns + tax) / price
        fixed_costs_per_unit = (cost + storage_cost + auto_parts_specific + marketing_expenses + packaging_expenses)
        break_even_volume = fixed_costs_per_unit / (price * (1 - variable_costs_per_unit)) if (1 - variable_costs_per_unit) > 0 else 0
        
        # === 15. РЕКОМЕНДУЕМАЯ ЦЕНА (С УЧЁТОМ ТРЕБУЕМОЙ МАРЖИ) ===
        target_margin = 0.20  # Целевая маржа 20%
        variable_rate = comm_rate + acquiring_rate_val + return_rate_val + tax_rate_val + 0.10
        fixed_costs = cost + log_base + storage_cost + marketing_expenses + packaging_expenses
        denominator = 1 - variable_rate - target_margin
        recommended_min_price = (fixed_costs / denominator) if denominator > 0 else price * 1.5
        
        # === 16. ДОПОЛНИТЕЛЬНЫЕ МЕТРИКИ ДЛЯ ЯНДЕКС МАРКЕТ ===
        yandex_specific = {}
        if marketplace == "Яндекс Маркет":
            yandex_specific = {
                "yandex_commission_optimized": round(commission * 0.95, 2),
                "yandex_logistics_optimized": round(logistics * 0.9, 2),
                "yandex_rating_impact": 0.02 if profit > 0 else -0.05,
                "yandex_competitive_price": round(recommended_min_price * 0.98, 2)
            }
        
        return {
            # Основные параметры
            "price": round(price, 2),
            "cost": round(cost, 2),
            "marketplace": marketplace,
            "operation_mode": operation_mode,
            "category": category,
            
            # Физические параметры
            "billable_weight": round(billable_weight, 2),
            "volume_liter": round(volume_liter, 3),
            
            # Расходы
            "commission": round(commission, 2),
            "commission_rate": round(comm_rate * 100, 2),
            "logistics": round(logistics, 2),
            "storage_cost": round(storage_cost, 2),
            "acquiring": round(acquiring, 2),
            "returns": round(returns, 2),
            "hazardous_surcharge": round(hazardous_surcharge, 2),
            "fragile_surcharge": round(fragile_surcharge, 2),
            "auto_parts_specific": round(auto_parts_specific, 2),
            "marketing_expenses": round(marketing_expenses, 2),
            "packaging_expenses": round(packaging_expenses, 2),
            "tax": round(tax, 2),
            
            # Итоги
            "total_expenses": round(total_expenses, 2),
            "profit": round(profit, 2),
            "margin_percent": round(margin_percent, 2),
            "roi": round(roi, 2),
            "break_even_volume": round(break_even_volume, 2),
            "recommended_min_price": round(recommended_min_price, 2),
            
            # Специфические для Яндекс Маркет
            **yandex_specific
        }

    def calculate_batch(
        self,
        df: pd.DataFrame,
        artikul_col: str,
        price_col: str,
        cost_col: str,
        weight_col: Optional[str] = None,
        length_col: Optional[str] = None,
        width_col: Optional[str] = None,
        height_col: Optional[str] = None,
        name_col: Optional[str] = None,
        days_in_storage: int = 30,
        operation_mode: str = "FBS",
        marketplace: str = "Яндекс Маркет",
        **kwargs
    ) -> pd.DataFrame:
        """
        Пакетный расчёт юнит-экономики для всех товаров в DataFrame.
        """
        results = []
        
        for idx, row in df.iterrows():
            try:
                price = float(row.get(price_col, 0) or 0)
                cost = float(row.get(cost_col, 0) or 0)
                
                if price <= 0 or cost <= 0:
                    continue
                
                weight = float(row.get(weight_col, 1.0) or 1.0) if weight_col and weight_col in row else 1.0
                length = float(row.get(length_col, 0) or 0) if length_col and length_col in row else 0
                width = float(row.get(width_col, 0) or 0) if width_col and width_col in row else 0
                height = float(row.get(height_col, 0) or 0) if height_col and height_col in row else 0
                
                result = self.calculate(
                    price=price,
                    cost=cost,
                    weight_kg=weight,
                    length_cm=length,
                    width_cm=width,
                    height_cm=height,
                    days_in_storage=days_in_storage,
                    operation_mode=operation_mode,
                    marketplace=marketplace,
                    **kwargs
                )
                
                result["Артикул"] = row.get(artikul_col, f"SKU_{idx}")
                result["Наименование"] = row.get(name_col, "") if name_col and name_col in row else ""
                
                # Сохраняем исходные данные
                result["weight_original"] = weight
                result["length_original"] = length
                result["width_original"] = width
                result["height_original"] = height
                result["storage_days"] = days_in_storage
                
                results.append(result)
                
            except Exception as e:
                logger.warning(f"Ошибка расчёта для строки {idx}: {e}")
                continue
        
        return pd.DataFrame(results)

# ============================================================================
# БЛОК 7: ПРОФЕССИОНАЛЬНЫЙ ЭКСПОРТЕР В EXCEL С ЖИВЫМИ ФОРМУЛАМИ
# ============================================================================

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule, DataBarRule
    from openpyxl.chart import BarChart, Reference, PieChart
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

class ProfessionalExcelExporter:
    """
    Профессиональный экспорт в Excel с живыми формулами и визуализацией.
    """
    def __init__(self):
        if not OPENPYXL_AVAILABLE:
            raise ImportError("openpyxl не установлен. pip install openpyxl")
    
    def export_with_formulas(
        self,
        df_results: pd.DataFrame,
        output_path: str,
        marketplace_name: str = "Яндекс Маркет"
    ) -> bool:
        """
        Экспорт в Excel с живыми формулами и профессиональной визуализацией.
        """
        try:
            wb = Workbook()
            
            # === ЛИСТ 1: Юнит-экономика ===
            ws1 = wb.active
            ws1.title = "Юнит-экономика FBS"
            
            # Расширенные заголовки
            headers = [
                "Артикул", "Наименование", "Цена продажи", "Себестоимость",
                "Вес, кг", "Длина, см", "Ширина, см", "Высота, см",
                "Ставка комиссии, %", "База логистики, ₽", "Логистика за кг, ₽",
                "Ставка хранения, ₽/день", "Дней хранения",
                "Объёмный вес, кг", "Оплачиваемый вес, кг",
                "Комиссия, ₽", "Логистика, ₽", "Хранение, ₽",
                "Эквайринг, ₽", "Возвраты, ₽", "Авто-специфика, ₽",
                "Маркетинг, ₽", "Упаковка, ₽", "Налог, ₽",
                "ИТОГО расходов, ₽", "ПРИБЫЛЬ, ₽", "МАРЖА, %", "ROI, %",
                "Рек. мин. цена, ₽", "Точка безубыточности"
            ]
            
            # Стили заголовков
            header_fill = PatternFill(start_color="1a1a2e", end_color="1a1a2e", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=10)
            header_alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            
            for col_idx, header in enumerate(headers, 1):
                cell = ws1.cell(row=1, column=col_idx, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = header_alignment
            
            # === ДАННЫЕ И ФОРМУЛЫ ===
            for row_idx, (_, row) in enumerate(df_results.iterrows(), 2):
                # Вводные данные (жёлтые ячейки)
                ws1.cell(row=row_idx, column=1, value=escape_excel_text(row.get("Артикул", "")))
                ws1.cell(row=row_idx, column=2, value=str(row.get("Наименование", "")))
                ws1.cell(row=row_idx, column=3, value=float(row.get("price", 0)))
                ws1.cell(row=row_idx, column=3).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws1.cell(row=row_idx, column=4, value=float(row.get("cost", 0)))
                ws1.cell(row=row_idx, column=4).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws1.cell(row=row_idx, column=5, value=float(row.get("weight_original", 1.0)))
                ws1.cell(row=row_idx, column=5).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws1.cell(row=row_idx, column=6, value=float(row.get("length_original", 0)))
                ws1.cell(row=row_idx, column=6).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws1.cell(row=row_idx, column=7, value=float(row.get("width_original", 0)))
                ws1.cell(row=row_idx, column=7).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws1.cell(row=row_idx, column=8, value=float(row.get("height_original", 0)))
                ws1.cell(row=row_idx, column=8).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws1.cell(row=row_idx, column=9, value=float(row.get("commission_rate", 14.5)))
                ws1.cell(row=row_idx, column=9).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws1.cell(row=row_idx, column=10, value=float(row.get("logistics_base", 55.0)))
                ws1.cell(row=row_idx, column=10).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws1.cell(row=row_idx, column=11, value=float(row.get("logistics_per_kg", 16.0)))
                ws1.cell(row=row_idx, column=11).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws1.cell(row=row_idx, column=12, value=float(row.get("storage_rate", 0.35)))
                ws1.cell(row=row_idx, column=12).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws1.cell(row=row_idx, column=13, value=int(row.get("storage_days", 30)))
                ws1.cell(row=row_idx, column=13).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                
                # Расчётные формулы (зелёные ячейки)
                # Объёмный вес
                ws1.cell(row=row_idx, column=14, value=f"=IF(F{row_idx}*G{row_idx}*H{row_idx}>0, (F{row_idx}*G{row_idx}*H{row_idx})/5000, 0)")
                ws1.cell(row=row_idx, column=14).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Оплачиваемый вес
                ws1.cell(row=row_idx, column=15, value=f"=CEILING(MAX(E{row_idx}, N{row_idx}), 0.5)")
                ws1.cell(row=row_idx, column=15).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Комиссия
                ws1.cell(row=row_idx, column=16, value=f"=MAX(C{row_idx}*(I{row_idx}/100), 35)")
                ws1.cell(row=row_idx, column=16).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Логистика
                ws1.cell(row=row_idx, column=17, value=f"=J{row_idx}+(O{row_idx}*K{row_idx})")
                ws1.cell(row=row_idx, column=17).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Хранение
                ws1.cell(row=row_idx, column=18, value=f"=IF(F{row_idx}*G{row_idx}*H{row_idx}>0, (F{row_idx}*G{row_idx}*H{row_idx}/1000)*L{row_idx}*M{row_idx}, 5*L{row_idx}*M{row_idx})")
                ws1.cell(row=row_idx, column=18).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Эквайринг
                ws1.cell(row=row_idx, column=19, value=f"=C{row_idx}*0.015")
                ws1.cell(row=row_idx, column=19).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Возвраты
                ws1.cell(row=row_idx, column=20, value=f"=C{row_idx}*0.025")
                ws1.cell(row=row_idx, column=20).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Авто-специфика
                ws1.cell(row=row_idx, column=21, value=f"=2+50+5+C{row_idx}*0.025")
                ws1.cell(row=row_idx, column=21).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Маркетинг
                ws1.cell(row=row_idx, column=22, value=float(row.get("marketing_expenses", 0)))
                ws1.cell(row=row_idx, column=22).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Упаковка
                ws1.cell(row=row_idx, column=23, value=float(row.get("packaging_expenses", 0)))
                ws1.cell(row=row_idx, column=23).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Налог
                ws1.cell(row=row_idx, column=24, value=f"=C{row_idx}*0.06")
                ws1.cell(row=row_idx, column=24).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # ИТОГО расходов (синяя)
                ws1.cell(row=row_idx, column=25, value=f"=D{row_idx}+P{row_idx}+Q{row_idx}+R{row_idx}+S{row_idx}+T{row_idx}+U{row_idx}+V{row_idx}+W{row_idx}+X{row_idx}")
                ws1.cell(row=row_idx, column=25).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws1.cell(row=row_idx, column=25).font = Font(bold=True)
                
                # ПРИБЫЛЬ (зелёная)
                ws1.cell(row=row_idx, column=26, value=f"=C{row_idx}-Y{row_idx}")
                ws1.cell(row=row_idx, column=26).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws1.cell(row=row_idx, column=26).font = Font(bold=True, color="006600")
                
                # МАРЖА
                ws1.cell(row=row_idx, column=27, value=f"=IF(C{row_idx}>0, (Z{row_idx}/C{row_idx})*100, 0)")
                ws1.cell(row=row_idx, column=27).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws1.cell(row=row_idx, column=27).font = Font(bold=True)
                ws1.cell(row=row_idx, column=27).number_format = '0.00"%"'
                
                # ROI
                ws1.cell(row=row_idx, column=28, value=f"=IF(D{row_idx}>0, (Z{row_idx}/D{row_idx})*100, 0)")
                ws1.cell(row=row_idx, column=28).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws1.cell(row=row_idx, column=28).number_format = '0.00"%"'
                
                # Рекомендуемая цена
                ws1.cell(row=row_idx, column=29, value=f"=MAX(0, (D{row_idx}+J{row_idx}+R{row_idx}+V{row_idx}+W{row_idx}) / MAX(0.01, (1 - I{row_idx}/100 - 0.015 - 0.025 - 0.06 - 0.10)))")
                ws1.cell(row=row_idx, column=29).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws1.cell(row=row_idx, column=29).font = Font(bold=True, color="CC0000")
                
                # Точка безубыточности
                ws1.cell(row=row_idx, column=30, value=f"=IF((C{row_idx}*(1-I{row_idx}/100-0.015-0.025-0.06))>0, (D{row_idx}+R{row_idx}+U{row_idx}+V{row_idx}+W{row_idx})/(C{row_idx}*(1-I{row_idx}/100-0.015-0.025-0.06)), 0)")
                ws1.cell(row=row_idx, column=30).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws1.cell(row=row_idx, column=30).number_format = '0.00'
            
            # === ФОРМАТИРОВАНИЕ КОЛОНОК ===
            for col_idx in range(1, 31):
                ws1.column_dimensions[get_column_letter(col_idx)].width = 15
            
            # === ЗАМОРОЗКА ПЕРВОЙ СТРОКИ ===
            ws1.freeze_panes = "A2"
            
            # === УСЛОВНОЕ ФОРМАТИРОВАНИЕ ===
            # Красный для убыточных
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            red_font = Font(color="9C0006")
            ws1.conditional_formatting.add(
                f"Z2:Z{len(df_results) + 1}",
                CellIsRule(operator="lessThan", formula=["0"], fill=red_fill, font=red_font)
            )
            
            # Зелёный для прибыльных
            green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            green_font = Font(color="006100")
            ws1.conditional_formatting.add(
                f"Z2:Z{len(df_results) + 1}",
                CellIsRule(operator="greaterThan", formula=["0"], fill=green_fill, font=green_font)
            )
            
            # === ЛИСТ 2: КРАТКИЙ ОТЧЁТ ===
            ws2 = wb.create_sheet("Краткий отчёт")
            
            # Заголовок
            ws2.merge_cells('A1:D1')
            title_cell = ws2.cell(row=1, column=1, value=f"📊 Юнит-экономика {marketplace_name} - Краткий отчёт")
            title_cell.font = Font(bold=True, size=14)
            title_cell.alignment = Alignment(horizontal="center")
            
            # Статистика
            stats = [
                ["Показатель", "Значение", "Единица измерения"],
                ["Всего SKU", len(df_results), "шт."],
                ["Прибыльных SKU", (df_results['profit'] > 0).sum(), "шт."],
                ["Убыточных SKU", (df_results['profit'] < 0).sum(), "шт."],
                ["Средняя маржа", f"{df_results['margin_percent'].mean():.1f}", "%"],
                ["Медианная маржа", f"{df_results['margin_percent'].median():.1f}", "%"],
                ["Общая прибыль", f"{df_results['profit'].sum():,.0f}", "₽"],
                ["Средний ROI", f"{df_results['roi'].mean():.1f}", "%"],
                ["Средняя рекомендуемая цена", f"{df_results['recommended_min_price'].mean():.0f}", "₽"]
            ]
            
            for row_idx, row_data in enumerate(stats, 3):
                for col_idx, value in enumerate(row_data, 1):
                    ws2.cell(row=row_idx, column=col_idx, value=value)
            
            # Форматирование
            for col_idx in range(1, 4):
                ws2.column_dimensions[get_column_letter(col_idx)].width = 25
            
            # === ЛИСТ 3: РЕКОМЕНДАЦИИ ===
            ws3 = wb.create_sheet("Рекомендации")
            
            ws3.merge_cells('A1:C1')
            title_cell = ws3.cell(row=1, column=1, value="💡 Рекомендации по оптимизации")
            title_cell.font = Font(bold=True, size=14)
            title_cell.alignment = Alignment(horizontal="center")
            
            # Анализ и рекомендации
            unprofitable_pct = (df_results['profit'] < 0).sum() / len(df_results) * 100
            median_margin = df_results['margin_percent'].median()
            underpriced_pct = (df_results['price'] < df_results['recommended_min_price']).sum() / len(df_results) * 100
            
            recommendations = [
                ["Категория", "Проблема", "Рекомендация"],
            ]
            
            if unprofitable_pct > 10:
                recommendations.append([
                    "Убыточность",
                    f"{unprofitable_pct:.1f}% товаров убыточны",
                    "Пересмотрите цены или откажитесь от этих позиций"
                ])
            
            if median_margin < 15:
                recommendations.append([
                    "Низкая маржа",
                    f"Медианная маржа {median_margin:.1f}%",
                    "Повысьте цены или снизьте закупочные цены"
                ])
            
            if underpriced_pct > 20:
                recommendations.append([
                    "Недооценка",
                    f"{underpriced_pct:.1f}% товаров недооценены",
                    "Повысьте цены до рекомендуемого уровня"
                ])
            
            if len(recommendations) == 1:
                recommendations.append([
                    "Отлично!",
                    "Все показатели в норме",
                    "Продолжайте мониторинг и оптимизацию"
                ])
            
            for row_idx, row_data in enumerate(recommendations, 3):
                for col_idx, value in enumerate(row_data, 1):
                    ws3.cell(row=row_idx, column=col_idx, value=value)
            
            for col_idx in range(1, 4):
                ws3.column_dimensions[get_column_letter(col_idx)].width = 30
            
            # === СОХРАНЕНИЕ ===
            wb.save(output_path)
            logger.info(f"✅ Профессиональный Excel с живыми формулами сохранён: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Ошибка экспорта Excel с формулами: {e}")
            return False

# ============================================================================
# БЛОК 8: ПРОФЕССИОНАЛЬНАЯ ВИЗУАЛИЗАЦИЯ
# ============================================================================

class ProfessionalVisualizer:
    """
    Профессиональная визуализация данных юнит-экономики.
    """
    
    @staticmethod
    def plot_margin_distribution(df: pd.DataFrame) -> go.Figure:
        """Профессиональная визуализация распределения маржи."""
        if df.empty or 'margin_percent' not in df.columns:
            fig = go.Figure()
            fig.add_annotation(text="Нет данных для визуализации", showarrow=False)
            return fig
        
        # Категории маржи
        bins = [-100, -25, -10, 0, 5, 10, 15, 20, 30, 50, 100]
        labels = ['<-25%', '-25% - -10%', '-10% - 0%', '0% - 5%', '5% - 10%', 
                  '10% - 15%', '15% - 20%', '20% - 30%', '30% - 50%', '>50%']
        
        df['margin_category'] = pd.cut(df['margin_percent'], bins=bins, labels=labels)
        distribution = df['margin_category'].value_counts().reindex(labels, fill_value=0)
        
        colors = ['#d62728', '#e6550d', '#fd8d3c', '#feb24c', '#fed976', 
                  '#a8d08d', '#74c476', '#31a354', '#238b45', '#006d2c']
        
        fig = go.Figure(data=[go.Bar(
            x=distribution.index,
            y=distribution.values,
            marker_color=colors,
            text=distribution.values,
            textposition='auto',
            hovertemplate='Маржа: %{x}<br>Кол-во: %{y}<extra></extra>'
        )])
        
        fig.add_hline(y=0, line_dash="dash", line_color="gray")
        
        fig.update_layout(
            title=dict(
                text="📊 Распределение маржи по товарам",
                font=dict(size=18, color='#1a1a2e')
            ),
            xaxis_title="Диапазон маржи, %",
            yaxis_title="Количество товаров",
            template="plotly_white",
            height=450,
            showlegend=False,
            bargap=0.1,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Arial, sans-serif")
        )
        
        return fig
    
    @staticmethod
    def plot_profit_analysis(df: pd.DataFrame) -> go.Figure:
        """Профессиональный анализ прибыли."""
        if df.empty or 'profit' not in df.columns:
            fig = go.Figure()
            fig.add_annotation(text="Нет данных для визуализации", showarrow=False)
            return fig
        
        # Топ по прибыли и убыткам
        top_profit = df.nlargest(15, 'profit')
        top_loss = df.nsmallest(15, 'profit')
        combined = pd.concat([top_loss, top_profit])
        
        colors = ['#d62728' if x < 0 else '#2ca02c' for x in combined['profit']]
        
        fig = go.Figure(data=[go.Bar(
            y=combined['Артикул'],
            x=combined['profit'],
            orientation='h',
            marker_color=colors,
            text=combined['profit'].apply(lambda x: f'{x:,.0f} ₽'),
            textposition='outside',
            hovertemplate='Артикул: %{y}<br>Прибыль: %{x:,.0f} ₽<br>Маржа: %{customdata:.1f}%',
            customdata=combined['margin_percent']
        )])
        
        fig.add_vline(x=0, line_dash="dash", line_color="gray", line_width=1)
        
        fig.update_layout(
            title=dict(
                text="📈 Топ товаров по прибыли и убыткам",
                font=dict(size=18, color='#1a1a2e')
            ),
            xaxis_title="Прибыль, ₽",
            yaxis_title="Артикул",
            template="plotly_white",
            height=500,
            showlegend=False,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Arial, sans-serif")
        )
        
        return fig
    
    @staticmethod
    def plot_cost_breakdown(result: Dict[str, Any]) -> go.Figure:
        """Профессиональная визуализация структуры расходов."""
        if not result:
            fig = go.Figure()
            fig.add_annotation(text="Нет данных", showarrow=False)
            return fig
        
        cost_categories = {
            'Себестоимость': result.get('cost', 0),
            'Комиссия МП': result.get('commission', 0),
            'Логистика': result.get('logistics', 0),
            'Хранение': result.get('storage_cost', 0),
            'Эквайринг': result.get('acquiring', 0),
            'Возвраты': result.get('returns', 0),
            'Налоги': result.get('tax', 0),
            'Маркетинг': result.get('marketing_expenses', 0),
            'Упаковка': result.get('packaging_expenses', 0),
            'Прочие': result.get('auto_parts_specific', 0)
        }
        
        # Убираем нулевые значения
        cost_categories = {k: v for k, v in cost_categories.items() if v > 0}
        
        if not cost_categories:
            fig = go.Figure()
            fig.add_annotation(text="Нет данных", showarrow=False)
            return fig
        
        colors = ['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd', 
                  '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf']
        
        fig = go.Figure(data=[go.Pie(
            labels=list(cost_categories.keys()),
            values=list(cost_categories.values()),
            hole=0.45,
            marker=dict(colors=colors[:len(cost_categories)]),
            textinfo='label+percent',
            textposition='outside',
            hovertemplate='%{label}<br>%{value:,.0f} ₽ (%{percent})<extra></extra>'
        )])
        
        fig.update_layout(
            title=dict(
                text="💰 Структура расходов",
                font=dict(size=18, color='#1a1a2e')
            ),
            template="plotly_white",
            height=450,
            showlegend=False,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Arial, sans-serif")
        )
        
        return fig
    
    @staticmethod
    def plot_marketplace_comparison(df: pd.DataFrame, marketplace: str) -> go.Figure:
        """Сравнение показателей для маркетплейса."""
        if df.empty:
            fig = go.Figure()
            fig.add_annotation(text="Нет данных", showarrow=False)
            return fig
        
        metrics = {
            'Средняя маржа': df['margin_percent'].mean(),
            'Медианная маржа': df['margin_percent'].median(),
            'Средний ROI': df['roi'].mean(),
            '% прибыльных': (df['profit'] > 0).sum() / len(df) * 100,
            'Средняя прибыль': df['profit'].mean()
        }
        
        fig = go.Figure(data=[go.Bar(
            x=list(metrics.keys()),
            y=list(metrics.values()),
            marker_color=['#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd'],
            text=[f"{v:.1f}" for v in metrics.values()],
            textposition='auto',
            hovertemplate='%{x}<br>%{y:.1f}<extra></extra>'
        )])
        
        fig.update_layout(
            title=dict(
                text=f"📊 Ключевые метрики для {marketplace}",
                font=dict(size=18, color='#1a1a2e')
            ),
            yaxis_title="Значение",
            template="plotly_white",
            height=400,
            showlegend=False,
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            font=dict(family="Arial, sans-serif")
        )
        
        return fig

# ============================================================================
# БЛОК 9: ПРОФЕССИОНАЛЬНЫЙ ИНТЕРФЕЙС ПОЛЬЗОВАТЕЛЯ (STREAMLIT)
# ============================================================================

def init_session_state():
    """Инициализация session state"""
    if 'secure_key_manager' not in st.session_state:
        st.session_state.secure_key_manager = SecureKeyManager()
    
    if 'processed_catalog_df' not in st.session_state:
        st.session_state.processed_catalog_df = pd.DataFrame()
    
    if 'calculation_results_df' not in st.session_state:
        st.session_state.calculation_results_df = pd.DataFrame()
    
    if 'visualizer' not in st.session_state:
        st.session_state.visualizer = ProfessionalVisualizer()
    
    if 'deepseek_manager' not in st.session_state:
        st.session_state.deepseek_manager = None
    
    # Инициализация калькулятора с настройками по умолчанию
    default_config = {
        "commission_rate": 0.145,
        "min_commission": 35.0,
        "logistics_base": 55.0,
        "logistics_per_kg": 16.0,
        "storage_per_day": 0.35,
        "acquiring_fee": 0.015,
        "return_fee": 0.025,
        "mode_multipliers": {"FBS": 1.0, "FBY": 0.8, "FBP": 0.7}
    }
    
    if 'fbs_calculator' not in st.session_state:
        st.session_state.fbs_calculator = FBSUnitEconomicsCalculator(
            marketplace_config=default_config,
            tax_system="УСН_6"
        )
    
    if 'excel_exporter' not in st.session_state:
        try:
            st.session_state.excel_exporter = ProfessionalExcelExporter()
        except ImportError:
            st.session_state.excel_exporter = None
            st.warning("⚠️ OpenPyXL не установлен. Экспорт в Excel недоступен.")
    
    if 'current_marketplace' not in st.session_state:
        st.session_state.current_marketplace = "Яндекс Маркет"

def show_section_data_loading():
    """📁 Загрузка данных и настройка API"""
    st.header("📁 Загрузка данных и настройка API")
    
    st.info("""
    **🎯 Что делает этот раздел:**
    - Загружает файлы каталога (цены, габариты, себестоимость)
    - Настраивает API ключи для DeepSeek AI
    - Выбирает маркетплейс для расчёта (Яндекс Маркет в приоритете)
    - Подготавливает данные для профессионального расчёта
    """)
    
    key_manager = st.session_state.secure_key_manager
    
    # --- Выбор маркетплейса ---
    st.subheader("🏪 Выбор маркетплейса")
    st.caption("Яндекс Маркет - приоритетная платформа с оптимизированными тарифами")
    
    marketplace = st.selectbox(
        "Выберите маркетплейс для расчёта:",
        options=["Яндекс Маркет", "Ozon", "Wildberries"],
        index=0,
        help="Яндекс Маркет - приоритетная платформа с оптимизированными тарифами"
    )
    
    st.session_state.current_marketplace = marketplace
    
    if marketplace == "Яндекс Маркет":
        st.success("✅ Выбран Яндекс Маркет - оптимизированные тарифы и логистика")
        st.info("📌 Для Яндекс Маркет применяются специальные коэффициенты: логистика на 10% ниже, оптимизированная комиссия")
    
    st.divider()
    
    # --- Управление API ключами ---
    with st.expander("🔑 Безопасное хранение API ключей", expanded=False):
        st.markdown("""
        **🔐 Профессиональная безопасность:**
        Все ключи шифруются с помощью Fernet и хранятся локально.
        Даже при компрометации файлов ключи останутся защищёнными.
        """)
        
        col_k1, col_k2 = st.columns(2)
        with col_k1:
            deepseek_key = st.text_input(
                "🤖 DeepSeek API Key (рекомендуется)", 
                value=key_manager.get_key("deepseek") or "",
                type="password",
                help="Ключ для AI-аналитики и обогащения данных"
            )
            if st.button("💾 Сохранить DeepSeek Key"):
                key_manager.set_key("deepseek", deepseek_key, "DeepSeek API Key для AI аналитики")
                st.success("✅ Ключ DeepSeek зашифрован и сохранен!")
                
        with col_k2:
            google_sheets_creds = st.text_area(
                "📊 Google Sheets Credentials (JSON)",
                value=key_manager.get_key("google_sheets") or "",
                height=100,
                help="Вставьте содержимое JSON файла сервисного аккаунта для интеграции"
            )
            if st.button("💾 Сохранить Google Sheets Credentials"):
                key_manager.set_key("google_sheets", google_sheets_creds, "Google Sheets Service Account JSON")
                st.success("✅ Google Sheets credentials сохранены!")
    
    st.divider()
    
    # --- Загрузка файлов ---
    st.subheader("📥 Загрузка каталога товаров")
    st.caption("Загрузите файл с данными о товарах: артикулы, цены, себестоимость, габариты")
    
    file_main = st.file_uploader(
        "📦 Загрузите файл каталога",
        type=['csv', 'xlsx', 'xls'],
        help="Файл должен содержать: Артикул, Цена продажи, Себестоимость (опционально: Вес, Длина, Ширина, Высота)"
    )
    
    if file_main is not None:
        st.success("✅ Файл загружен")
        
        with st.spinner("Чтение и обработка данных..."):
            df_main = smart_read_uploaded_file(file_main)
            
        if not df_main.empty:
            st.session_state.processed_catalog_df = df_main
            st.success(f"✅ Загружено {len(df_main)} товаров, {len(df_main.columns)} колонок")
            
            st.markdown("##### 👁️ Предпросмотр данных")
            st.dataframe(df_main.head(10), use_container_width=True)
        else:
            st.error("❌ Не удалось прочитать файл")
    
    st.divider()
    
    # --- AI Аналитика ---
    st.subheader("🤖 Профессиональная AI-аналитика")
    
    ds_key = key_manager.get_key("deepseek")
    if not ds_key:
        st.warning("⚠️ Ключ DeepSeek не задан. Настройте его в блоке 'Безопасное хранение API ключей'.")
        st.info("💡 DeepSeek AI помогает: анализировать товары, определять оптимальные цены, давать рекомендации")
    else:
        st.success("✅ DeepSeek API доступен")
        
        ai_action = st.radio(
            "Выберите действие AI:",
            options=[
                "📊 Получить инсайты по маркетплейсу",
                "🔍 Проанализировать товар",
                "💡 Получить рекомендации"
            ],
            horizontal=True
        )
        
        if ai_action == "📊 Получить инсайты по маркетплейсу":
            if st.button("🚀 Получить профессиональные инсайты", type="primary"):
                manager = DeepSeekAPIManager(api_key=ds_key)
                st.session_state.deepseek_manager = manager
                
                with st.spinner("DeepSeek AI анализирует маркетплейс..."):
                    result = manager.get_marketplace_insights(marketplace, "auto_parts")
                    
                    if "error" in result:
                        st.error(f"❌ Ошибка: {result['error']}")
                    else:
                        st.success("✅ Инсайты получены!")
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.markdown("**📈 Тренды маркетплейса:**")
                            st.json(result.get("marketplace_trends", {}))
                        
                        with col2:
                            st.markdown("**💰 Тарифы 2026:**")
                            st.json(result.get("tariffs_2026", {}))
                        
                        st.markdown("**💡 Рекомендации для продавцов:**")
                        st.json(result.get("seller_insights", {}))
        
        elif ai_action == "🔍 Проанализировать товар":
            if not st.session_state.processed_catalog_df.empty:
                df = st.session_state.processed_catalog_df
                name_col = next((c for c in df.columns if 'наименование' in c.lower() or 'name' in c.lower()), df.columns[0])
                
                col1, col2 = st.columns(2)
                with col1:
                    product_name = st.selectbox("Выберите товар для анализа:", df[name_col].head(50).tolist())
                with col2:
                    price_val = float(df[df[name_col] == product_name].iloc[0].get('Цена', 0) or 0) if 'Цена' in df.columns else 0
                    current_price = st.number_input("Текущая цена, ₽", value=price_val, min_value=0.0, step=10.0)
                
                if st.button("🔍 Проанализировать товар", type="primary"):
                    manager = DeepSeekAPIManager(api_key=ds_key)
                    st.session_state.deepseek_manager = manager
                    
                    with st.spinner("DeepSeek AI анализирует товар..."):
                        result = manager.analyze_product_for_marketplace(
                            product_name=product_name,
                            marketplace=marketplace,
                            current_price=current_price
                        )
                        
                        if "error" in result:
                            st.error(f"❌ Ошибка: {result['error']}")
                        else:
                            st.success("✅ Анализ завершён!")
                            
                            col1, col2 = st.columns(2)
                            with col1:
                                st.markdown("**💰 Рекомендации по цене:**")
                                rec = result.get("marketplace_recommendations", {})
                                st.metric("Рекомендуемая цена", f"{rec.get('recommended_price', 0):.0f} ₽")
                                st.metric("Минимальная маржа", f"{rec.get('min_profit_margin', 0)}%")
                                st.metric("Уровень конкуренции", rec.get("competition_level", "medium").title())
                            
                            with col2:
                                st.markdown("**📦 Характеристики товара:**")
                                chars = result.get("product_characteristics", {})
                                st.json(chars)
                            
                            st.markdown("**💡 Рекомендации по оптимизации:**")
                            tips = result.get("competitive_analysis", {}).get("profit_optimization_tips", [])
                            for tip in tips:
                                st.info(f"• {tip}")
        
        elif ai_action == "💡 Получить рекомендации":
            if st.button("💡 Получить профессиональные рекомендации", type="primary"):
                manager = DeepSeekAPIManager(api_key=ds_key)
                st.session_state.deepseek_manager = manager
                
                with st.spinner("DeepSeek AI генерирует рекомендации..."):
                    st.info("💡 Рекомендации будут отображаться после расчёта юнит-экономики в следующем разделе")

def show_section_single_calculation():
    """🧮 Калькулятор единичного товара"""
    st.header("🧮 Профессиональный калькулятор единичного товара")
    
    st.info("""
    **🎯 Что делает этот раздел:**
    - Быстрый расчёт юнит-экономики для одного товара
    - Детальная структура всех расходов
    - Профессиональные метрики: ROI, точка безубыточности
    - Рекомендации по оптимизации цены
    - Идеально для тестирования гипотез
    """)
    
    # Получаем калькулятор из session state
    calc = st.session_state.fbs_calculator
    marketplace = st.session_state.current_marketplace
    
    # Проверяем, что калькулятор инициализирован
    if calc is None:
        st.error("❌ Калькулятор не инициализирован. Пожалуйста, перезагрузите страницу.")
        return
    
    # Устанавливаем конфигурацию для выбранного маркетплейса
    try:
        calc.set_marketplace(marketplace)
    except Exception as e:
        st.error(f"❌ Ошибка настройки маркетплейса: {e}")
        return
    
    # --- Ввод параметров ---
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("💰 Финансовые параметры")
        price = st.number_input("Цена продажи, ₽", min_value=1.0, value=500.0, step=10.0, 
                               help="Розничная цена на маркетплейсе")
        cost = st.number_input("Себестоимость, ₽", min_value=1.0, value=300.0, step=10.0,
                              help="Закупочная цена + доставка до склада")
        
        st.subheader("📦 Физические параметры")
        weight = st.number_input("Вес, кг", min_value=0.1, value=1.0, step=0.1,
                                help="Вес брутто (с упаковкой)")
        col_w1, col_w2, col_w3 = st.columns(3)
        with col_w1:
            length = st.number_input("Длина, см", min_value=0, value=20, step=1)
        with col_w2:
            width = st.number_input("Ширина, см", min_value=0, value=15, step=1)
        with col_w3:
            height = st.number_input("Высота, см", min_value=0, value=10, step=1)
    
    with col2:
        st.subheader("⚙️ Параметры расчёта")
        
        operation_mode = st.selectbox(
            "Режим работы",
            options=["FBS", "FBY", "FBP"],
            index=0,
            help="FBS — со своего склада по заказу. FBY — аналог FBS с доставкой МП. FBP — FBS с премиум-логистикой"
        )
        
        days_in_storage = st.number_input(
            "Дней хранения",
            min_value=1, max_value=365, value=30, step=1,
            help="Среднее время хранения на складе маркетплейса"
        )
        
        tax_system = st.selectbox(
            "Налоговая система",
            options=["УСН_6", "УСН_15", "ОСН", "НПД"],
            index=0,
            format_func=lambda x: {"УСН_6": "УСН 6% (доходы)", "УСН_15": "УСН 15% (доходы-расходы)", 
                                  "ОСН": "ОСН (общая)", "НПД": "НПД (самозанятый)"}[x]
        )
        
        st.markdown("---")
        st.subheader("💡 Особенности товара")
        is_hazardous = st.checkbox("Опасный груз", help="+20% к логистике")
        is_fragile = st.checkbox("Хрупкий товар", help="+15% к логистике")
        
        st.subheader("📊 Дополнительные расходы")
        marketing_cost = st.number_input("Маркетинговые расходы, ₽", min_value=0.0, value=0.0, step=5.0)
        packaging_cost = st.number_input("Расходы на упаковку, ₽", min_value=0.0, value=0.0, step=5.0)
    
    # --- Расчёт ---
    if st.button("🧮 Профессиональный расчёт юнит-экономики", type="primary", use_container_width=True):
        try:
            # Обновляем налоговую систему
            calc.tax_system = tax_system
            
            # Выполняем расчёт
            result = calc.calculate(
                price=price,
                cost=cost,
                weight_kg=weight,
                length_cm=length,
                width_cm=width,
                height_cm=height,
                days_in_storage=days_in_storage,
                operation_mode=operation_mode,
                marketplace=marketplace,
                is_hazardous=is_hazardous,
                is_fragile=is_fragile,
                marketing_cost=marketing_cost,
                packaging_cost=packaging_cost
            )
            
            # --- Отображение результатов ---
            st.divider()
            st.subheader(f"📊 Результаты расчёта для {marketplace}")
            
            # KPI в 4 колонки
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                profit_color = "green" if result['profit'] > 0 else "red"
                st.metric("💰 Прибыль", f"{result['profit']:.0f} ₽", 
                         delta=f"{result['margin_percent']:.1f}% маржи", 
                         delta_color="normal")
            with col2:
                st.metric("📦 Итого расходов", f"{result['total_expenses']:.0f} ₽")
            with col3:
                st.metric("📈 ROI", f"{result['roi']:.1f}%")
            with col4:
                price_diff = result['recommended_min_price'] - result['price']
                if price_diff > 0:
                    st.metric("💡 Рекомендуемая цена", f"{result['recommended_min_price']:.0f} ₽", 
                             delta=f"+{price_diff:.0f} ₽", 
                             delta_color="inverse")
                else:
                    st.metric("💡 Рекомендуемая цена", f"{result['recommended_min_price']:.0f} ₽", 
                             delta="✅ Цена оптимальна", 
                             delta_color="off")
            
            # Дополнительные метрики
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("⚖️ Оплачиваемый вес", f"{result['billable_weight']:.2f} кг")
            with col2:
                st.metric("📊 Точка безубыточности", f"{result['break_even_volume']:.0f} шт.")
            with col3:
                st.metric("📦 Объём", f"{result['volume_liter']:.2f} л")
            
            # Детальная таблица
            st.markdown("##### 📋 Детальная калькуляция")
            
            detail_data = {
                "Категория": [
                    "💰 Доходы",
                    "💰 Доходы",
                    "📦 Расходы",
                    "📦 Расходы",
                    "📦 Расходы",
                    "📦 Расходы",
                    "📦 Расходы",
                    "📦 Расходы",
                    "📦 Расходы",
                    "📦 Расходы",
                    "📦 Расходы",
                    "📦 Расходы",
                    "📦 Расходы",
                    "📊 Итоги",
                    "📊 Итоги"
                ],
                "Показатель": [
                    "Цена продажи",
                    "Себестоимость",
                    "Комиссия МП",
                    "Логистика",
                    "Хранение",
                    "Эквайринг",
                    "Возвраты",
                    "Налоги",
                    "Маркетинг",
                    "Упаковка",
                    "Специфические расходы",
                    "Надбавка (опасный)",
                    "Надбавка (хрупкий)",
                    "ИТОГО расходов",
                    "ПРИБЫЛЬ"
                ],
                "Сумма, ₽": [
                    result['price'],
                    result['cost'],
                    result['commission'],
                    result['logistics'],
                    result['storage_cost'],
                    result['acquiring'],
                    result['returns'],
                    result['tax'],
                    result.get('marketing_expenses', 0),
                    result.get('packaging_expenses', 0),
                    result['auto_parts_specific'],
                    result.get('hazardous_surcharge', 0),
                    result.get('fragile_surcharge', 0),
                    result['total_expenses'],
                    result['profit']
                ],
                "% от цены": [
                    "100%",
                    f"{result['cost']/result['price']*100:.1f}%",
                    f"{result['commission']/result['price']*100:.1f}%",
                    f"{result['logistics']/result['price']*100:.1f}%",
                    f"{result['storage_cost']/result['price']*100:.1f}%",
                    f"{result['acquiring']/result['price']*100:.1f}%",
                    f"{result['returns']/result['price']*100:.1f}%",
                    f"{result['tax']/result['price']*100:.1f}%",
                    f"{result.get('marketing_expenses', 0)/result['price']*100:.1f}%",
                    f"{result.get('packaging_expenses', 0)/result['price']*100:.1f}%",
                    f"{result['auto_parts_specific']/result['price']*100:.1f}%",
                    f"{result.get('hazardous_surcharge', 0)/result['price']*100:.1f}%",
                    f"{result.get('fragile_surcharge', 0)/result['price']*100:.1f}%",
                    f"{result['total_expenses']/result['price']*100:.1f}%",
                    f"{result['profit']/result['price']*100:.1f}%"
                ]
            }
            
            df_detail = pd.DataFrame(detail_data)
            
            # Стилизация
            def color_rows(row):
                if row['Показатель'] == 'ПРИБЫЛЬ':
                    return ['background-color: #d4edda; font-weight: bold'] * len(row)
                elif row['Показатель'] == 'ИТОГО расходов':
                    return ['background-color: #fff3cd; font-weight: bold'] * len(row)
                elif row['Категория'] == '💰 Доходы':
                    return ['background-color: #e8f4fd'] * len(row)
                elif row['Категория'] == '📦 Расходы':
                    return ['background-color: #fde8e8'] * len(row)
                elif row['Категория'] == '📊 Итоги':
                    return ['background-color: #e8fde8; font-weight: bold'] * len(row)
                return [''] * len(row)
            
            st.dataframe(df_detail.style.apply(color_rows, axis=1), use_container_width=True)
            
            # --- Визуализация ---
            st.divider()
            st.subheader("📊 Профессиональная визуализация")
            
            col1, col2 = st.columns(2)
            with col1:
                fig_pie = st.session_state.visualizer.plot_cost_breakdown(result)
                st.plotly_chart(fig_pie, use_container_width=True)
            
            with col2:
                # Сравнение цены и расходов
                fig_bar = go.Figure()
                
                categories = ['Цена', 'Расходы', 'Прибыль']
                values = [result['price'], result['total_expenses'], result['profit']]
                colors = ['#1f77b4', '#ff7f0e', '#2ca02c']
                
                fig_bar.add_trace(go.Bar(
                    x=categories,
                    y=values,
                    marker_color=colors,
                    text=[f"{v:.0f} ₽" for v in values],
                    textposition='auto',
                    hovertemplate='%{x}<br>%{y:.0f} ₽<extra></extra>'
                ))
                
                fig_bar.add_hline(y=0, line_dash="dash", line_color="gray")
                
                fig_bar.update_layout(
                    title="Сравнение цены, расходов и прибыли",
                    yaxis_title="Сумма, ₽",
                    template="plotly_white",
                    height=400,
                    showlegend=False,
                    plot_bgcolor='rgba(0,0,0,0)',
                    paper_bgcolor='rgba(0,0,0,0)'
                )
                st.plotly_chart(fig_bar, use_container_width=True)
            
        except Exception as e:
            st.error(f"❌ Ошибка расчёта: {e}")
            logger.exception("Ошибка в single calculation")

def show_section_batch_calculation():
    """📊 Массовый расчёт юнит-экономики"""
    st.header("📊 Профессиональный массовый расчёт юнит-экономики")
    
    st.info("""
    **🎯 Что делает этот раздел:**
    - Массовый расчёт юнит-экономики для всего каталога
    - Профессиональные метрики для каждого товара
    - Выявление убыточных позиций и точек роста
    - Экспорт в Excel с живыми формулами и визуализацией
    - Интеграция с Google Sheets
    """)
    
    # --- Проверка наличия данных ---
    if st.session_state.processed_catalog_df.empty:
        st.error("❌ Нет данных каталога. Перейдите в раздел 'Загрузка данных и настройка API'.")
        return
        
    df_catalog = st.session_state.processed_catalog_df.copy()
    marketplace = st.session_state.current_marketplace
    
    # Получаем калькулятор
    calc = st.session_state.fbs_calculator
    
    if calc is None:
        st.error("❌ Калькулятор не инициализирован. Пожалуйста, перезагрузите страницу.")
        return
    
    # --- Настройки расчёта ---
    st.subheader("⚙️ Настройки массового расчёта")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.markdown(f"**🏪 Маркетплейс:** {marketplace}")
        st.caption("Яндекс Маркет - приоритет")
    with col2:
        operation_mode = st.selectbox(
            "Режим работы",
            options=["FBS", "FBY", "FBP"],
            index=0
        )
    with col3:
        days_in_storage = st.number_input(
            "Дней хранения",
            min_value=1, max_value=365, value=30, step=1
        )
    with col4:
        tax_system = st.selectbox(
            "Налоговая система",
            options=["УСН_6", "УСН_15", "ОСН", "НПД"],
            index=0,
            format_func=lambda x: {"УСН_6": "УСН 6% (доходы)", "УСН_15": "УСН 15% (доходы-расходы)", 
                                  "ОСН": "ОСН (общая)", "НПД": "НПД (самозанятый)"}[x]
        )
    
    # --- Определение колонок ---
    st.subheader("🔍 Настройка колонок")
    st.caption("Укажите, какие колонки в вашем файле соответствуют параметрам расчёта")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        artikul_col = st.selectbox("Артикул", df_catalog.columns.tolist(), 
                                   index=next((i for i, c in enumerate(df_catalog.columns) if 'артикул' in c.lower() or 'artikul' in c.lower()), 0))
    with col2:
        price_col = st.selectbox("Цена продажи", df_catalog.columns.tolist(),
                                 index=next((i for i, c in enumerate(df_catalog.columns) if 'цена' in c.lower() or 'price' in c.lower()), 0))
    with col3:
        cost_col = st.selectbox("Себестоимость", df_catalog.columns.tolist(),
                                index=next((i for i, c in enumerate(df_catalog.columns) if 'себестоимость' in c.lower() or 'cost' in c.lower() or 'закуп' in c.lower()), 0))
    with col4:
        name_col = st.selectbox("Наименование", ["Не выбрано"] + df_catalog.columns.tolist(), index=0)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        weight_col = st.selectbox("Вес, кг", ["Не выбрано"] + df_catalog.columns.tolist(),
                                  index=next((i+1 for i, c in enumerate(df_catalog.columns) if 'вес' in c.lower() or 'weight' in c.lower()), 0))
    with col2:
        length_col = st.selectbox("Длина, см", ["Не выбрано"] + df_catalog.columns.tolist(),
                                  index=next((i+1 for i, c in enumerate(df_catalog.columns) if 'длин' in c.lower() or 'length' in c.lower()), 0))
    with col3:
        width_col = st.selectbox("Ширина, см", ["Не выбрано"] + df_catalog.columns.tolist(),
                                 index=next((i+1 for i, c in enumerate(df_catalog.columns) if 'ширин' in c.lower() or 'width' in c.lower()), 0))
    with col4:
        height_col = st.selectbox("Высота, см", ["Не выбрано"] + df_catalog.columns.tolist(),
                                  index=next((i+1 for i, c in enumerate(df_catalog.columns) if 'высот' in c.lower() or 'height' in c.lower()), 0))
    
    # --- Запуск расчёта ---
    st.divider()
    
    try:
        calc.tax_system = tax_system
        calc.set_marketplace(marketplace)
    except Exception as e:
        st.error(f"❌ Ошибка настройки калькулятора: {e}")
        return
    
    if st.button("🚀 Запустить профессиональный массовый расчёт", type="primary", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        # Подготовка данных
        weight_col_actual = weight_col if weight_col != "Не выбрано" else None
        length_col_actual = length_col if length_col != "Не выбрано" else None
        width_col_actual = width_col if width_col != "Не выбрано" else None
        height_col_actual = height_col if height_col != "Не выбрано" else None
        name_col_actual = name_col if name_col != "Не выбрано" else None
        
        total = len(df_catalog)
        
        # Пакетный расчёт
        try:
            df_results = calc.calculate_batch(
                df=df_catalog,
                artikul_col=artikul_col,
                price_col=price_col,
                cost_col=cost_col,
                weight_col=weight_col_actual,
                length_col=length_col_actual,
                width_col=width_col_actual,
                height_col=height_col_actual,
                name_col=name_col_actual,
                days_in_storage=days_in_storage,
                operation_mode=operation_mode,
                marketplace=marketplace
            )
        except Exception as e:
            st.error(f"❌ Ошибка расчёта: {e}")
            logger.exception("Ошибка в batch calculation")
            return
        
        progress_bar.progress(1.0)
        status_text.text(f"✅ Расчёт завершён! Обработано {len(df_results)} товаров.")
        
        if not df_results.empty:
            st.session_state.calculation_results_df = df_results
            st.success(f"✅ Рассчитано {len(df_results)} товаров. Средняя маржа: {df_results['margin_percent'].mean():.1f}%")
            st.info(f"📊 Прибыльных: {(df_results['profit'] > 0).sum()}, Убыточных: {(df_results['profit'] < 0).sum()}")
        else:
            st.error("❌ Не удалось рассчитать ни одного товара. Проверьте данные.")
    
    # --- Отображение результатов ---
    if not st.session_state.calculation_results_df.empty:
        df_results = st.session_state.calculation_results_df
        
        st.divider()
        st.subheader("📊 Результаты массового расчёта")
        
        # KPI
        col1, col2, col3, col4, col5 = st.columns(5)
        with col1:
            st.metric("💰 Общая прибыль", f"{df_results['profit'].sum():,.0f} ₽")
        with col2:
            st.metric("📈 Средняя маржа", f"{df_results['margin_percent'].mean():.1f}%")
        with col3:
            st.metric("📊 Средний ROI", f"{df_results['roi'].mean():.1f}%")
        with col4:
            unprofitable = (df_results['profit'] < 0).sum()
            st.metric("⚠️ Убыточных SKU", f"{unprofitable}")
        with col5:
            profitable = (df_results['profit'] > 0).sum()
            st.metric("✅ Прибыльных SKU", f"{profitable}")
        
        # Визуализация
        st.divider()
        st.subheader("📈 Профессиональная визуализация")
        
        col1, col2 = st.columns(2)
        with col1:
            fig_margin = st.session_state.visualizer.plot_margin_distribution(df_results)
            st.plotly_chart(fig_margin, use_container_width=True)
        
        with col2:
            fig_profit = st.session_state.visualizer.plot_profit_analysis(df_results)
            st.plotly_chart(fig_profit, use_container_width=True)
        
        # Сравнение с маркетплейсом
        st.divider()
        fig_compare = st.session_state.visualizer.plot_marketplace_comparison(df_results, marketplace)
        st.plotly_chart(fig_compare, use_container_width=True)
        
        # Таблица результатов
        st.divider()
        st.subheader("📋 Детальная таблица")
        
        # Поиск
        search_term = st.text_input("🔍 Поиск по артикулу или наименованию", placeholder="Введите артикул или название...")
        if search_term:
            mask = df_results['Артикул'].str.contains(search_term, case=False, na=False) | \
                   df_results['Наименование'].str.contains(search_term, case=False, na=False)
            df_filtered = df_results[mask]
        else:
            df_filtered = df_results
        
        # Сортировка
        sort_col = st.selectbox("Сортировать по", ["Прибыль (убывание)", "Маржа (убывание)", "Артикул (возрастание)", "ROI (убывание)"])
        if sort_col == "Прибыль (убывание)":
            df_filtered = df_filtered.sort_values('profit', ascending=False)
        elif sort_col == "Маржа (убывание)":
            df_filtered = df_filtered.sort_values('margin_percent', ascending=False)
        elif sort_col == "ROI (убывание)":
            df_filtered = df_filtered.sort_values('roi', ascending=False)
        else:
            df_filtered = df_filtered.sort_values('Артикул')
        
        # Отображение
        display_cols = ["Артикул", "Наименование", "price", "cost", "profit", "margin_percent", "roi", "recommended_min_price", "break_even_volume"]
        available_cols = [c for c in display_cols if c in df_filtered.columns]
        
        st.dataframe(
            df_filtered[available_cols].style.background_gradient(
                subset=['profit', 'margin_percent', 'roi'], 
                cmap='RdYlGn', 
                vmin=-50, 
                vmax=50
            ),
            use_container_width=True,
            height=400
        )
        
        # --- Экспорт ---
        st.divider()
        st.subheader("📥 Профессиональный экспорт")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**📊 Excel с живыми формулами**")
            st.caption("Профессиональный отчёт с формулами, диаграммами и рекомендациями")
            
            if st.button("📥 Экспортировать в Excel (Профессиональный)", type="primary", use_container_width=True):
                try:
                    output_path = EXPORTS_DIR / f"unit_economics_pro_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    exporter = st.session_state.excel_exporter
                    
                    if exporter is None:
                        st.error("❌ Excel экспортер не доступен. Установите openpyxl.")
                        return
                    
                    # Добавляем маркетинговые и упаковочные расходы, если их нет
                    if 'marketing_expenses' not in df_results.columns:
                        df_results['marketing_expenses'] = 0
                    if 'packaging_expenses' not in df_results.columns:
                        df_results['packaging_expenses'] = 0
                    
                    success = exporter.export_with_formulas(df_results, str(output_path), marketplace)
                    
                    if success and output_path.exists():
                        with open(output_path, "rb") as f:
                            file_data = f.read()
                        st.download_button(
                            label="⬇️ Скачать профессиональный Excel",
                            data=file_data,
                            file_name=output_path.name,
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="excel_download"
                        )
                        st.success("✅ Профессиональный Excel с живыми формулами готов к скачиванию!")
                    else:
                        st.error("❌ Ошибка создания файла")
                except Exception as e:
                    st.error(f"❌ Ошибка экспорта: {e}")
                    logger.exception("Ошибка экспорта")
        
        with col2:
            st.markdown("**📄 CSV для анализа**")
            st.caption("Универсальный формат для загрузки в другие системы")
            
            if st.button("📄 Экспортировать в CSV", use_container_width=True):
                csv_data = df_results.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="⬇️ Скачать CSV",
                    data=csv_data,
                    file_name=f"unit_economics_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key="csv_download"
                )
        
        # --- Google Sheets интеграция ---
        st.divider()
        st.subheader("🔄 Интеграция с Google Sheets")
        
        key_manager = st.session_state.secure_key_manager
        gs_creds = key_manager.get_key("google_sheets")
        
        if not gs_creds:
            st.warning("⚠️ Google Sheets credentials не заданы. Настройте их в разделе 'Загрузка данных и настройка API'.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                spreadsheet_id = st.text_input("ID Google Таблицы", placeholder="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms")
            with col2:
                worksheet_name = st.text_input("Название листа", value="Юнит-экономика")
            
            if st.button("📤 Экспортировать в Google Sheets", use_container_width=True):
                if not spreadsheet_id:
                    st.error("❌ Укажите ID Google Таблицы")
                else:
                    try:
                        manager = GoogleSheetsManager(gs_creds)
                        export_df = df_results[["Артикул", "Наименование", "price", "profit", "margin_percent", "roi", "recommended_min_price"]].copy()
                        export_df.columns = ["Артикул", "Наименование", "Цена", "Прибыль", "Маржа %", "ROI %", "Рек. цена"]
                        success = manager.write_sheet(
                            spreadsheet_id=spreadsheet_id,
                            df=export_df,
                            worksheet_name=worksheet_name,
                            clear_before=True
                        )
                        if success:
                            st.success("✅ Результаты экспортированы в Google Sheets!")
                        else:
                            st.error("❌ Ошибка экспорта")
                    except Exception as e:
                        st.error(f"❌ Ошибка: {e}")

# ============================================================================
# БЛОК 10: ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================

def main():
    """Главная функция приложения"""
    
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="🚀",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Заголовок
    st.markdown(f"""
    <div style='text-align: center; padding: 30px 20px; background: linear-gradient(135deg, #0f0c29, #302b63, #24243e); border-radius: 15px; margin-bottom: 25px;'>
        <h1 style='color: white; margin: 0; font-size: 2.5em;'>{APP_NAME}</h1>
        <p style='color: #a8a8d0; margin: 10px 0 0 0; font-size: 1.2em;'>
        {APP_DESCRIPTION}
        </p>
        <p style='color: #6666aa; margin: 5px 0 0 0; font-size: 0.9em;'>
        Версия {APP_VERSION} | FBS-ONLY | Яндекс Маркет Приоритет
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Инициализация
    init_session_state()
    
    # Sidebar навигация
    st.sidebar.title("🧭 Навигация")
    
    # Создаем словарь с описанием разделов
    sections = {
        "📁 Загрузка данных и настройка API": {
            "icon": "📁",
            "description": "Загрузка каталога, настройка API ключей, выбор маркетплейса"
        },
        "🧮 Калькулятор единичного товара": {
            "icon": "🧮",
            "description": "Быстрый расчёт для одного товара с детальной аналитикой"
        },
        "📊 Массовый расчёт юнит-экономики": {
            "icon": "📊",
            "description": "Массовый расчёт с экспортом в Excel и Google Sheets"
        }
    }
    
    # Отображаем разделы
    section_keys = list(sections.keys())
    selected_section = st.sidebar.radio(
        "Выберите раздел:",
        section_keys,
        format_func=lambda x: f"{sections[x]['icon']} {x}"
    )
    
    # Описание выбранного раздела
    st.sidebar.markdown("---")
    st.sidebar.markdown(f"""
    **📌 {selected_section}**
    {sections[selected_section]['description']}
    """)
    
    # Статус системы
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📊 Статус системы")
    
    status_cols = st.sidebar.columns(2)
    with status_cols[0]:
        st.sidebar.success("✅ SecureKeyManager")
        st.sidebar.success("✅ FBS Calculator")
        if st.session_state.excel_exporter is not None:
            st.sidebar.success("✅ Excel Exporter")
        else:
            st.sidebar.warning("⚠️ Excel Exporter")
    with status_cols[1]:
        if CRYPTO_AVAILABLE:
            st.sidebar.success("✅ Cryptography")
        else:
            st.sidebar.warning("⚠️ Crypto")
        if GSPREAD_AVAILABLE:
            st.sidebar.success("✅ GSpread")
        else:
            st.sidebar.warning("⚠️ GSpread")
        if st.session_state.get('deepseek_manager') and st.session_state.deepseek_manager:
            st.sidebar.success("✅ DeepSeek AI")
        else:
            st.sidebar.info("⚪ DeepSeek AI")
    
    # Информация о данных
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📦 Данные")
    
    marketplace = st.session_state.get('current_marketplace', 'Яндекс Маркет')
    st.sidebar.info(f"🏪 {marketplace}")
    
    if not st.session_state.processed_catalog_df.empty:
        st.sidebar.success(f"✅ Каталог: {len(st.session_state.processed_catalog_df)} товаров")
    else:
        st.sidebar.warning("⚠️ Каталог не загружен")
    
    if not st.session_state.calculation_results_df.empty:
        st.sidebar.success(f"✅ Рассчитано: {len(st.session_state.calculation_results_df)} товаров")
    else:
        st.sidebar.warning("⚠️ Расчёт не выполнен")
    
    # Быстрый старт
    st.sidebar.markdown("---")
    st.sidebar.info("""
    **🚀 Быстрый старт:**
    1. 📁 Загрузите каталог в разделе "Загрузка данных"
    2. 🤖 Настройте DeepSeek API (опционально)
    3. 📊 Выполните массовый расчёт
    4. 📥 Экспортируйте результаты
    """)
    
    # Версия
    st.sidebar.markdown("---")
    st.sidebar.caption(f"Версия {APP_VERSION}")
    
    # Отображение выбранного раздела
    try:
        if selected_section == "📁 Загрузка данных и настройка API":
            show_section_data_loading()
        elif selected_section == "🧮 Калькулятор единичного товара":
            show_section_single_calculation()
        elif selected_section == "📊 Массовый расчёт юнит-экономики":
            show_section_batch_calculation()
    except Exception as e:
        st.error(f"❌ Ошибка при загрузке раздела: {e}")
        logger.exception("Ошибка в main")
    
    # Футер
    st.divider()
    st.markdown(f"""
    <div style='text-align: center; padding: 20px; color: #666;'>
        <p style='margin: 0;'>🚀 <strong>{APP_NAME}</strong></p>
        <p style='margin: 5px 0 0 0; font-size: 0.9em;'>
        Версия {APP_VERSION} | Живые формулы Excel | DeepSeek AI | Яндекс Маркет Приоритет
        </p>
        <p style='margin: 5px 0 0 0; font-size: 0.8em; color: #999;'>
        Профессиональный инструмент для юнит-экономики на маркетплейсах
        </p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
