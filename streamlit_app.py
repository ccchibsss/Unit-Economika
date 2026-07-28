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
warnings.filterwarnings('ignore')

# ============================================================================
# БЛОК 0: БАЗОВАЯ КОНФИГУРАЦИЯ И ИМПОРТЫ
# ============================================================================

# === Версия приложения ===
APP_VERSION = "2.0.0"
APP_NAME = "🚀 FBS Юнит-экономика PRO 2026"
APP_DESCRIPTION = "Сквозная юнит-экономика с живыми формулами Excel, интеграцией DeepSeek API и Google Sheets"

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
# БЛОК 3: КРОСС-СВЯЗЫВАНИЕ ДАННЫХ (OE-номера)
# ============================================================================

class OECrossLinker:
    """
    Класс для умного кросс-связывания данных через ОЕ-номера.
    """
    
    @staticmethod
    def split_oe_numbers(oe_string: str) -> List[str]:
        """Разделяет строку ОЕ-номеров по точке с запятой и очищает"""
        if pd.isna(oe_string) or not oe_string:
            return []
        return [oe.strip() for oe in str(oe_string).split(';') if oe.strip()]

    @staticmethod
    def normalize_oe(oe: str) -> str:
        """Нормализует ОЕ-номер для сравнения"""
        if pd.isna(oe) or not oe:
            return ""
        return re.sub(r'[^0-9A-Za-z]', '', str(oe).upper())

    @staticmethod
    def normalize_artikul(art: str) -> str:
        """Нормализует артикул для сравнения"""
        if pd.isna(art) or not art:
            return ""
        return re.sub(r'[^0-9A-Za-z]', '', str(art).upper())

    def link_and_fill_missing(
        self,
        df_main: pd.DataFrame,
        df_oe: pd.DataFrame,
        main_art_col: str,
        main_oe_col: str,
        oe_art_col: str,
        oe_oe_col: str,
        cols_to_fill: List[str]
    ) -> pd.DataFrame:
        """
        Связывает основной DataFrame с DataFrame ОЕ и заполняет пропуски.
        """
        if df_main.empty or df_oe.empty:
            return df_main.copy()

        result = df_main.copy()
        
        # Создаем нормализованные ключи
        result['_norm_art'] = result[main_art_col].apply(self.normalize_artikul)
        result['_norm_oe'] = result[main_oe_col].apply(self.normalize_oe)
        
        df_oe_work = df_oe.copy()
        df_oe_work['_norm_art'] = df_oe_work[oe_art_col].apply(self.normalize_artikul)
        df_oe_work['_norm_oe_raw'] = df_oe_work[oe_oe_col].apply(self.normalize_oe)
        
        # Разбиваем ОЕ номера через ';'
        df_oe_work['_oe_list'] = df_oe_work[oe_oe_col].apply(self.split_oe_numbers)
        df_oe_work['_norm_oe_list'] = df_oe_work['_oe_list'].apply(
            lambda lst: [self.normalize_oe(oe) for oe in lst]
        )
        
        # Explode
        df_exploded = df_oe_work.explode('_norm_oe_list').rename(columns={'_norm_oe_list': '_norm_oe'})
        df_exploded = df_exploded[df_exploded['_norm_oe'] != ""]
        df_exploded = df_exploded.drop_duplicates(subset=['_norm_oe'], keep='first')
        
        # Merge по ОЕ
        merge_cols = ['_norm_oe'] + cols_to_fill
        existing_cols = [c for c in merge_cols if c in df_exploded.columns]
        if existing_cols:
            merged = result.merge(
                df_exploded[existing_cols],
                on='_norm_oe',
                how='left',
                suffixes=('', '_from_oe')
            )
        else:
            merged = result
        
        # Заполняем пропуски
        filled_count = 0
        for col in cols_to_fill:
            if col not in result.columns:
                continue
            oe_col = f"{col}_from_oe"
            if oe_col in merged.columns:
                mask = merged[col].isna() | (merged[col] == 0) | (merged[col] == '')
                merged.loc[mask, col] = merged.loc[mask, oe_col]
                filled_count += int(mask.sum())
                merged = merged.drop(columns=[oe_col])

        logger.info(f"✅ OECrossLinker: Заполнено {filled_count} пропусков через ОЕ-связывание.")
        
        # Очистка временных колонок
        final_cols = [c for c in merged.columns if not c.startswith('_norm_') and not c.startswith('_oe_') and not c.startswith('_')]
        
        return merged[final_cols]

    def build_cross_references_column(
        self,
        df_main: pd.DataFrame,
        df_cross: pd.DataFrame,
        main_art_col: str,
        cross_art_col: str,
        cross_analog_col: str,
        output_col_name: str = "Кроссы (аналоги)"
    ) -> pd.DataFrame:
        """
        Добавляет в основной DataFrame колонку с аналогами (кроссами).
        """
        if df_main.empty or df_cross.empty:
            result = df_main.copy()
            result[output_col_name] = ""
            return result

        result = df_main.copy()
        result['_norm_art'] = result[main_art_col].apply(self.normalize_artikul)
        
        df_cross_work = df_cross.copy()
        df_cross_work['_norm_art'] = df_cross_work[cross_art_col].apply(self.normalize_artikul)
        
        # Группируем аналоги
        analogs_grouped = (
            df_cross_work.groupby('_norm_art')[cross_analog_col]
            .apply(lambda x: '; '.join([str(v).strip() for v in x if pd.notna(v) and str(v).strip()]))
            .reset_index()
            .rename(columns={cross_analog_col: output_col_name})
        )
        
        result = result.merge(analogs_grouped, on='_norm_art', how='left')
        result[output_col_name] = result[output_col_name].fillna("")
        result = result.drop(columns=['_norm_art'])
        
        return result

# ============================================================================
# БЛОК 4: DEEPSEEK API ИНТЕГРАЦИЯ
# ============================================================================

class DeepSeekAPIManager:
    """
    Менеджер для работы с DeepSeek API.
    Поддерживает два режима: "Обогащение каталога" или "Актуализация тарифов".
    """
    def __init__(self, api_key: str, base_url: str = "https://api.deepseek.com/v1"):
        self.api_key = api_key
        self.base_url = base_url
        
    def is_available(self) -> bool:
        return bool(self.api_key)
        
    def enrich_catalog(self, product_name: str, current_category: str = "") -> Dict[str, Any]:
        """
        Режим 1: Обогащение каталога.
        """
        if not self.is_available():
            return {"error": "DeepSeek API недоступен или ключ не задан"}
            
        prompt = f"""
        Ты эксперт по автозапчастям. Проанализируй название товара: "{product_name}".
        Текущая категория (если есть): "{current_category}".
        
        Верни строго JSON в следующем формате:
        {{
            "parent_category": "Автозапчасти",
            "group_category": "Например: Подвеска, Двигатель, Тормозная система",
            "subgroup_category": "Например: Сайлентблоки, Поршни, Колодки",
            "product_type": "Например: SUSPENSION, ENGINE, BRAKE",
            "hazardous": false,
            "fragile": false,
            "confidence": 0.95
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
                timeout=30
            )
            
            if response.status_code == 200:
                content = response.json()["choices"][0]["message"]["content"]
                return json.loads(content)
            else:
                return {"error": f"HTTP {response.status_code}: {response.text}"}
                
        except Exception as e:
            logger.error(f"Ошибка DeepSeek (обогащение): {e}")
            return {"error": str(e)}
            
    def update_tariffs(self, marketplace: str, category: str = "auto_parts") -> Dict[str, Any]:
        """
        Режим 2: Актуализация тарифов маркетплейсов.
        """
        if not self.is_available():
            return {"error": "DeepSeek API недоступен или ключ не задан"}
            
        prompt = f"""
        Ты аналитик данных маркетплейсов. Предоставь актуальные тарифы для "{marketplace}" 
        для категории "{category}" на 2026 год.
        
        Верни строго JSON в следующем формате (числа как float, без процентов):
        {{
            "commission_rate": 0.15,
            "min_commission": 30.0,
            "logistics_base": 50.0,
            "logistics_per_kg": 15.0,
            "storage_per_day": 0.3,
            "return_fee": 0.02,
            "acquiring_fee": 0.015,
            "source": "DeepSeek AI Estimate"
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
                timeout=30
            )
            
            if response.status_code == 200:
                content = response.json()["choices"][0]["message"]["content"]
                return json.loads(content)
            else:
                return {"error": f"HTTP {response.status_code}: {response.text}"}
                
        except Exception as e:
            logger.error(f"Ошибка DeepSeek (тарифы): {e}")
            return {"error": str(e)}

# ============================================================================
# БЛОК 5: GOOGLE SHEETS ИНТЕГРАЦИЯ
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
# БЛОК 6: ГЕНЕРАТОР ЖИВЫХ ФОРМУЛ EXCEL
# ============================================================================

class ExcelFormulaBuilder:
    """
    Генератор живых формул Excel для юнит-экономики.
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

    def build_total_expenses_formula(self, row: int = 2) -> str:
        cost = self._get_cell("cost", row)
        commission = self._get_cell("commission", row)
        logistics = self._get_cell("logistics", row)
        storage = self._get_cell("storage", row)
        acquiring = self._get_cell("acquiring", row)
        tax = self._get_cell("tax", row)
        returns = self._get_cell("returns", row)
        return f"={cost}+{commission}+{logistics}+{storage}+{acquiring}+{tax}+{returns}"

    def build_profit_formula(self, row: int = 2) -> str:
        price = self._get_cell("price", row)
        expenses = self._get_cell("total_expenses", row)
        return f"={price}-{expenses}"

    def build_margin_formula(self, row: int = 2) -> str:
        profit = self._get_cell("profit", row)
        price = self._get_cell("price", row)
        return f"=IF({price}>0, ({profit}/{price})*100, 0)"

    def build_recommended_price_formula(self, row: int = 2) -> str:
        cost = self._get_cell("cost", row)
        logistics_base = self._get_cell("logistics_base", row)
        storage = self._get_cell("storage", row)
        comm_rate = self._get_cell("commission_rate", row)
        acquiring_rate = self._get_cell("acquiring_rate", row)
        tax_rate = self._get_cell("tax_rate", row)
        return f"=MAX(0, ({cost}+{logistics_base}+{storage}) / MAX(0.01, (1 - {comm_rate} - {acquiring_rate} - {tax_rate} - 0.10)))"

# ============================================================================
# БЛОК 7: FBS-ONLY РАСЧЁТ ЮНИТ-ЭКОНОМИКИ
# ============================================================================

class FBSUnitEconomicsCalculator:
    """
    Калькулятор юнит-экономики для режима FBS.
    """
    
    ALLOWED_MODES = ["FBS", "FBY"]
    
    def __init__(self, marketplace_config: Dict[str, Any], tax_system: str = "УСН_6"):
        self.config = marketplace_config
        self.tax_system = tax_system
        
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
        category: str = "auto_parts",
        is_hazardous: bool = False,
        is_fragile: bool = False,
        commission_rate: Optional[float] = None,
        logistics_base: Optional[float] = None,
        logistics_per_kg: Optional[float] = None,
        storage_rate: Optional[float] = None,
        acquiring_rate: Optional[float] = None,
        return_rate: Optional[float] = None,
        tax_rate: Optional[float] = None
    ) -> Dict[str, Any]:
        """
        Основной метод расчёта юнит-экономики для FBS.
        """
        if operation_mode not in self.ALLOWED_MODES:
            raise ValueError(f"Режим {operation_mode} не поддерживается. Используйте FBS или FBY.")
        
        if price <= 0 or cost <= 0:
            raise ValueError("Цена и себестоимость должны быть положительными")
        
        # === 1. КОМИССИЯ МП ===
        comm_rate = commission_rate if commission_rate is not None else self.config.get("commission_rate", 0.15)
        category_rates = self.config.get("category_rates", {})
        if category in category_rates:
            comm_rate = category_rates[category]
            
        commission = max(price * comm_rate, self.config.get("min_commission", 0.0))
        
        # === 2. ЛОГИСТИКА FBS ===
        vol_weight = (length_cm * width_cm * height_cm) / 5000.0 if length_cm > 0 else 0
        billable_weight = max(weight_kg, vol_weight)
        billable_weight = math.ceil(billable_weight * 2) / 2
        
        log_base = logistics_base if logistics_base is not None else self.config.get("logistics_base", 50.0)
        log_per_kg = logistics_per_kg if logistics_per_kg is not None else self.config.get("logistics_per_kg", 15.0)
        logistics = log_base + (billable_weight * log_per_kg)
        
        mode_multipliers = self.config.get("mode_multipliers", {"FBS": 1.0, "FBY": 0.75})
        logistics *= mode_multipliers.get(operation_mode, 1.0)
        
        # === 3. ХРАНЕНИЕ ===
        volume_liter = (length_cm * width_cm * height_cm) / 1000.0 if length_cm > 0 else 5.0
        storage_rate_val = storage_rate if storage_rate is not None else self.config.get("storage_per_day", 0.3)
        
        if days_in_storage <= 60:
            storage_multiplier = 1.0
        elif days_in_storage <= 90:
            storage_multiplier = 2.0
        elif days_in_storage <= 180:
            storage_multiplier = 4.0
        else:
            storage_multiplier = 8.0
            
        storage_cost = volume_liter * storage_rate_val * days_in_storage * storage_multiplier
        
        # === 4. ЭКВАЙРИНГ ===
        acquiring_rate_val = acquiring_rate if acquiring_rate is not None else self.config.get("acquiring_fee", 0.015)
        acquiring = price * acquiring_rate_val
        
        # === 5. ВОЗВРАТЫ ===
        return_rate_val = return_rate if return_rate is not None else self.config.get("return_fee", 0.02)
        returns = price * return_rate_val
        
        # === 6. НАДБАВКИ ===
        hazardous_surcharge = price * 0.02 if is_hazardous else 0.0
        fragile_surcharge = price * 0.01 if is_fragile else 0.0
        
        # === 7. СПЕЦИФИЧЕСКИЕ РАСХОДЫ ===
        auto_parts_specific = 1.5 + 45.0 + 3.0 + price * 0.02
        
        # === 8. НАЛОГ ===
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
            profit_before_tax = price - cost - commission - logistics - storage_cost - acquiring - returns - auto_parts_specific
            tax = max(0, profit_before_tax * tax_rate_val)
            if self.tax_system == "УСН_15":
                min_tax = price * tax_config.get("min_rate", 0.01)
                tax = max(tax, min_tax)
            
        # === 9. ИТОГО РАСХОДОВ ===
        total_expenses = (
            cost + commission + logistics + storage_cost + acquiring + returns +
            hazardous_surcharge + fragile_surcharge + auto_parts_specific + tax
        )
        
        # === 10. ПРИБЫЛЬ И МЕТРИКИ ===
        profit = price - total_expenses
        margin_percent = (profit / price * 100) if price > 0 else 0
        roi = (profit / cost * 100) if cost > 0 else 0
        
        # === 11. РЕКОМЕНДУЕМАЯ ЦЕНА ===
        variable_rate = comm_rate + acquiring_rate_val + return_rate_val + tax_rate_val + 0.10
        fixed_costs = cost + log_base + storage_cost
        denominator = 1 - variable_rate
        recommended_min_price = (fixed_costs / denominator) if denominator > 0 else 0
        
        return {
            "price": round(price, 2),
            "cost": round(cost, 2),
            "operation_mode": operation_mode,
            "billable_weight": round(billable_weight, 2),
            "volume_liter": round(volume_liter, 3),
            "commission": round(commission, 2),
            "commission_rate": round(comm_rate * 100, 2),
            "logistics": round(logistics, 2),
            "storage_cost": round(storage_cost, 2),
            "acquiring": round(acquiring, 2),
            "returns": round(returns, 2),
            "hazardous_surcharge": round(hazardous_surcharge, 2),
            "fragile_surcharge": round(fragile_surcharge, 2),
            "auto_parts_specific": round(auto_parts_specific, 2),
            "tax": round(tax, 2),
            "total_expenses": round(total_expenses, 2),
            "profit": round(profit, 2),
            "margin_percent": round(margin_percent, 2),
            "roi": round(roi, 2),
            "recommended_min_price": round(recommended_min_price, 2)
        }

# ============================================================================
# БЛОК 8: ЭКСПОРТЕР С ЖИВЫМИ ФОРМУЛАМИ EXCEL
# ============================================================================

try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.formatting.rule import CellIsRule, ColorScaleRule
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

class LiveExcelExporter:
    """
    Экспорт результатов расчёта в Excel с ЖИВЫМИ ФОРМУЛАМИ.
    """
    def __init__(self):
        if not OPENPYXL_AVAILABLE:
            raise ImportError("openpyxl не установлен. pip install openpyxl")
            
    def export_with_formulas(
        self,
        df_results: pd.DataFrame,
        output_path: str,
        marketplace_name: str = "Ozon"
    ) -> bool:
        """
        Экспорт DataFrame в Excel с живыми формулами.
        """
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = "Юнит-экономика FBS"
            
            # === ЗАГОЛОВКИ ===
            headers = [
                "Артикул", "Наименование", "Цена продажи", "Себестоимость",
                "Вес, кг", "Длина, см", "Ширина, см", "Высота, см",
                "Ставка комиссии, %", "База логистики, ₽", "Логистика за кг, ₽",
                "Ставка хранения, ₽/день", "Дней хранения",
                "Объёмный вес, кг", "Оплачиваемый вес, кг",
                "Комиссия, ₽", "Логистика, ₽", "Хранение, ₽",
                "Эквайринг, ₽", "Возвраты, ₽", "Авто-специфика, ₽", "Налог, ₽",
                "ИТОГО расходов, ₽", "ПРИБЫЛЬ, ₽", "МАРЖА, %", "Рек. мин. цена, ₽"
            ]
            
            for col_idx, header in enumerate(headers, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
                cell.alignment = Alignment(horizontal="center", vertical="center")
                
            # === ДАННЫЕ И ФОРМУЛЫ ===
            for row_idx, (_, row) in enumerate(df_results.iterrows(), 2):
                # Вводные данные (желтые ячейки)
                ws.cell(row=row_idx, column=1, value=escape_excel_text(row.get("Артикул", "")))
                ws.cell(row=row_idx, column=2, value=str(row.get("Наименование", "")))
                ws.cell(row=row_idx, column=3, value=float(row.get("price", 0)))
                ws.cell(row=row_idx, column=3).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=4, value=float(row.get("cost", 0)))
                ws.cell(row=row_idx, column=4).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=5, value=float(row.get("weight", 1.0)))
                ws.cell(row=row_idx, column=5).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=6, value=float(row.get("length", 0)))
                ws.cell(row=row_idx, column=6).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=7, value=float(row.get("width", 0)))
                ws.cell(row=row_idx, column=7).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=8, value=float(row.get("height", 0)))
                ws.cell(row=row_idx, column=8).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=9, value=float(row.get("commission_rate", 15.0)))
                ws.cell(row=row_idx, column=9).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=10, value=float(row.get("logistics_base", 50.0)))
                ws.cell(row=row_idx, column=10).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=11, value=float(row.get("logistics_per_kg", 15.0)))
                ws.cell(row=row_idx, column=11).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=12, value=float(row.get("storage_rate", 0.3)))
                ws.cell(row=row_idx, column=12).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                ws.cell(row=row_idx, column=13, value=int(row.get("storage_days", 30)))
                ws.cell(row=row_idx, column=13).fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")
                
                # Формулы (зеленые ячейки)
                # Объёмный вес = (F*G*H)/5000
                ws.cell(row=row_idx, column=14, value=f"=IF(F{row_idx}*G{row_idx}*H{row_idx}>0, (F{row_idx}*G{row_idx}*H{row_idx})/5000, 0)")
                ws.cell(row=row_idx, column=14).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Оплачиваемый вес = MAX(E, N)
                ws.cell(row=row_idx, column=15, value=f"=CEILING(MAX(E{row_idx}, N{row_idx}), 0.5)")
                ws.cell(row=row_idx, column=15).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Комиссия = MAX(C * (I/100), 30)
                ws.cell(row=row_idx, column=16, value=f"=MAX(C{row_idx}*(I{row_idx}/100), 30)")
                ws.cell(row=row_idx, column=16).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Логистика = J + (O * K)
                ws.cell(row=row_idx, column=17, value=f"=J{row_idx}+(O{row_idx}*K{row_idx})")
                ws.cell(row=row_idx, column=17).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Хранение = (F*G*H/1000) * L * M
                ws.cell(row=row_idx, column=18, value=f"=IF(F{row_idx}*G{row_idx}*H{row_idx}>0, (F{row_idx}*G{row_idx}*H{row_idx}/1000)*L{row_idx}*M{row_idx}, 5*L{row_idx}*M{row_idx})")
                ws.cell(row=row_idx, column=18).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Эквайринг = C * 0.015
                ws.cell(row=row_idx, column=19, value=f"=C{row_idx}*0.015")
                ws.cell(row=row_idx, column=19).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Возвраты = C * 0.02
                ws.cell(row=row_idx, column=20, value=f"=C{row_idx}*0.02")
                ws.cell(row=row_idx, column=20).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Авто-специфика = 1.5 + 45 + 3 + C*0.02
                ws.cell(row=row_idx, column=21, value=f"=1.5+45+3+C{row_idx}*0.02")
                ws.cell(row=row_idx, column=21).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # Налог = C * 0.06 (УСН 6%)
                ws.cell(row=row_idx, column=22, value=f"=C{row_idx}*0.06")
                ws.cell(row=row_idx, column=22).fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
                
                # ИТОГО расходов = D + P + Q + R + S + T + U + V
                ws.cell(row=row_idx, column=23, value=f"=D{row_idx}+P{row_idx}+Q{row_idx}+R{row_idx}+S{row_idx}+T{row_idx}+U{row_idx}+V{row_idx}")
                ws.cell(row=row_idx, column=23).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws.cell(row=row_idx, column=23).font = Font(bold=True)
                
                # ПРИБЫЛЬ = C - W
                ws.cell(row=row_idx, column=24, value=f"=C{row_idx}-W{row_idx}")
                ws.cell(row=row_idx, column=24).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws.cell(row=row_idx, column=24).font = Font(bold=True, color="006600")
                
                # МАРЖА = (X / C) * 100
                ws.cell(row=row_idx, column=25, value=f"=IF(C{row_idx}>0, (X{row_idx}/C{row_idx})*100, 0)")
                ws.cell(row=row_idx, column=25).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws.cell(row=row_idx, column=25).font = Font(bold=True)
                ws.cell(row=row_idx, column=25).number_format = '0.00"%"'
                
                # Рек. мин. цена = (D + J + R) / (1 - I/100 - 0.015 - 0.02 - 0.06 - 0.10)
                ws.cell(row=row_idx, column=26, value=f"=MAX(0, (D{row_idx}+J{row_idx}+R{row_idx}) / MAX(0.01, (1 - I{row_idx}/100 - 0.015 - 0.02 - 0.06 - 0.10)))")
                ws.cell(row=row_idx, column=26).fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
                ws.cell(row=row_idx, column=26).font = Font(bold=True, color="CC0000")
                
            # === АВТОШИРИНА КОЛОНОК ===
            for col_idx in range(1, 27):
                ws.column_dimensions[get_column_letter(col_idx)].width = 15
                
            # === ЗАМОРОЗКА ПЕРВОЙ СТРОКИ ===
            ws.freeze_panes = "A2"
            
            # === УСЛОВНОЕ ФОРМАТИРОВАНИЕ ===
            red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            red_font = Font(color="9C0006")
            ws.conditional_formatting.add(
                f"X2:X{len(df_results) + 1}",
                CellIsRule(operator="lessThan", formula=["0"], fill=red_fill, font=red_font)
            )
            
            # === СОХРАНЕНИЕ ===
            wb.save(output_path)
            logger.info(f"✅ Excel с живыми формулами сохранён: {output_path}")
            return True
            
        except Exception as e:
            logger.error(f"Ошибка экспорта Excel с формулами: {e}")
            return False

# ============================================================================
# БЛОК 9: ИНТЕРФЕЙС ПОЛЬЗОВАТЕЛЯ (STREAMLIT)
# ============================================================================

def init_session_state():
    """Инициализация session state"""
    if 'secure_key_manager' not in st.session_state:
        st.session_state.secure_key_manager = SecureKeyManager()
    
    if 'cross_processor' not in st.session_state:
        st.session_state.cross_processor = OECrossLinker()
    
    if 'processed_catalog_df' not in st.session_state:
        st.session_state.processed_catalog_df = pd.DataFrame()
    
    if 'calculation_results_df' not in st.session_state:
        st.session_state.calculation_results_df = pd.DataFrame()
    
    if 'fbs_calculator' not in st.session_state:
        default_config = {
            "commission_rate": 0.15,
            "min_commission": 30.0,
            "logistics_base": 50.0,
            "logistics_per_kg": 15.0,
            "storage_per_day": 0.3,
            "acquiring_fee": 0.015,
            "return_fee": 0.02,
            "mode_multipliers": {"FBS": 1.0, "FBY": 0.75},
            "category_rates": {}
        }
        st.session_state.fbs_calculator = FBSUnitEconomicsCalculator(
            marketplace_config=default_config,
            tax_system="УСН_6"
        )
    
    if 'deepseek_manager' not in st.session_state:
        st.session_state.deepseek_manager = None

def escape_excel_text(value: Any) -> str:
    """Экранирует строку для Excel, чтобы предотвратить автоматическое преобразование в дату или формулу."""
    if pd.isna(value) or value is None:
        return ""
    
    s = str(value).strip()
    if not s:
        return s
    
    # Проверка на формулы
    if s.startswith(('=', '+', '-', '@')):
        return f"'{s}"
    
    # Проверка на потенциальные даты (например, "1-2", "OCT", "2023-10", "10/12")
    if re.match(r'^\d+[-/]\d+([-/]\d+)?$', s) or re.match(r'^[A-Za-z]{3,4}[-/]\d+$', s, re.IGNORECASE):
        return f"'{s}"
    
    # Проверка на артикулы типа "12345-678" или "A123-B45"
    if re.match(r'^[A-Za-z0-9]+[-][A-Za-z0-9]+$', s):
        return f"'{s}"
        
    return s

def show_section1_data_loading():
    """Раздел 1: Загрузка и связывание данных"""
    st.header("📁 Раздел 1: Загрузка и связывание данных")
    
    st.info("""
    **В этом разделе:**
    1. Загрузите файлы каталога (Габариты, ОЕ, Кроссы)
    2. Настройте кросс-связывание для заполнения пропусков
    3. Сохраните API ключи в зашифрованном виде
    4. Выберите режим работы DeepSeek AI
    """)
    
    key_manager = st.session_state.secure_key_manager
    
    # --- Управление API ключами ---
    with st.expander("🔑 Управление API ключами (Безопасное хранение)", expanded=False):
        st.markdown("Ключи шифруются и сохраняются локально.")
        
        col_k1, col_k2, col_k3 = st.columns(3)
        with col_k1:
            deepseek_key = st.text_input(
                "DeepSeek API Key", 
                value=key_manager.get_key("deepseek") or "",
                type="password",
                help="Ключ для обогащения каталога или актуализации тарифов"
            )
            if st.button("💾 Сохранить DeepSeek Key"):
                key_manager.set_key("deepseek", deepseek_key, "DeepSeek API Key для AI функций")
                st.success("✅ Ключ DeepSeek зашифрован и сохранен!")
                
        with col_k2:
            ozon_key = st.text_input(
                "Ozon API Key (опционально)", 
                value=key_manager.get_key("ozon") or "",
                type="password"
            )
            if st.button("💾 Сохранить Ozon Key"):
                key_manager.set_key("ozon", ozon_key, "Ozon Seller API Key")
                st.success("✅ Ключ Ozon зашифрован и сохранен!")
                
        with col_k3:
            google_sheets_creds = st.text_area(
                "Google Sheets Credentials (JSON)",
                value=key_manager.get_key("google_sheets") or "",
                height=100,
                help="Вставьте содержимое JSON файла сервисного аккаунта"
            )
            if st.button("💾 Сохранить Google Sheets Credentials"):
                key_manager.set_key("google_sheets", google_sheets_creds, "Google Sheets Service Account JSON")
                st.success("✅ Google Sheets credentials сохранены!")
    
    st.divider()
    
    # --- Загрузка файлов ---
    st.subheader("📥 Загрузка файлов каталога")
    col_f1, col_f2, col_f3 = st.columns(3)
    
    with col_f1:
        file_main = st.file_uploader(
            "📦 Основной файл (Габариты/Цены)",
            type=['csv', 'xlsx', 'xls'],
            key="upload_main",
            help="Должен содержать Артикул, Цену, Себестоимость"
        )
        
    with col_f2:
        file_oe = st.file_uploader(
            "🔧 Файл ОЕ номеров (опционально)",
            type=['csv', 'xlsx', 'xls'],
            key="upload_oe",
            help="Используется для заполнения пропусков (вес, габариты)"
        )
        
    with col_f3:
        file_cross = st.file_uploader(
            "🔗 Файл Кроссов/Аналогов (опционально)",
            type=['csv', 'xlsx', 'xls'],
            key="upload_cross",
            help="Используется для создания столбца 'Кроссы (аналоги)'"
        )
    
    # --- Обработка и связывание ---
    if file_main is not None:
        st.success("✅ Основной файл загружен.")
        
        with st.spinner("Чтение и очистка данных..."):
            df_main = smart_read_uploaded_file(file_main)
            df_oe = smart_read_uploaded_file(file_oe) if file_oe else None
            df_cross = smart_read_uploaded_file(file_cross) if file_cross else None
            
        if not df_main.empty:
            st.subheader("⚙️ Настройка кросс-связывания")
            
            # Авто-детект колонок
            art_cols = [c for c in df_main.columns if 'артикул' in c.lower() or 'artikul' in c.lower() or 'sku' in c.lower()]
            main_art_col = art_cols[0] if art_cols else df_main.columns[0]
            
            oe_cols = [c for c in df_main.columns if 'ое' in c.lower() or 'oe' in c.lower()]
            main_oe_col = oe_cols[0] if oe_cols else ""
            
            col_m1, col_m2 = st.columns(2)
            with col_m1:
                main_art_col = st.selectbox("Колонка Артикула (Основной)", df_main.columns.tolist(), 
                                            index=df_main.columns.tolist().index(main_art_col) if main_art_col in df_main.columns else 0)
                main_oe_col = st.selectbox("Колонка ОЕ номера (Основной)", ["Не выбрано"] + df_main.columns.tolist(), 
                                           index=df_main.columns.tolist().index(main_oe_col) + 1 if main_oe_col in df_main.columns else 0)
                
            with col_m2:
                if df_oe is not None and not df_oe.empty:
                    oe_art_col = st.selectbox("Колонка Артикула (Файл ОЕ)", df_oe.columns.tolist(), index=0)
                    oe_oe_col = st.selectbox("Колонка ОЕ номера (Файл ОЕ)", df_oe.columns.tolist(), 
                                             index=df_oe.columns.tolist().index(next((c for c in df_oe.columns if 'ое' in c.lower() or 'oe' in c.lower()), df_oe.columns[0])) if any('ое' in c.lower() or 'oe' in c.lower() for c in df_oe.columns) else 0)
                else:
                    oe_art_col = "Артикул"
                    oe_oe_col = "ОЕ номер"
                    
                if df_cross is not None and not df_cross.empty:
                    cross_art_col = st.selectbox("Колонка Артикула (Файл Кроссов)", df_cross.columns.tolist(), index=0)
                    cross_analog_col = st.selectbox("Колонка Аналога (Файл Кроссов)", df_cross.columns.tolist(), 
                                                    index=df_cross.columns.tolist().index(next((c for c in df_cross.columns if 'аналог' in c.lower() or 'analog' in c.lower()), df_cross.columns[0])) if any('аналог' in c.lower() or 'analog' in c.lower() for c in df_cross.columns) else 0)
                else:
                    cross_art_col = "Артикул"
                    cross_analog_col = "Аналог"
                
            if st.button("🚀 Обработать и связать данные", type="primary"):
                with st.spinner("Выполняется кросс-связывание..."):
                    try:
                        processor = st.session_state.cross_processor
                        
                        actual_oe_col = main_oe_col if main_oe_col != "Не выбрано" else "ОЕ номер"
                        df_result = processor.link_and_fill_missing(
                            df_main=df_main,
                            df_oe=df_oe,
                            main_art_col=main_art_col,
                            main_oe_col=actual_oe_col,
                            oe_art_col=oe_art_col,
                            oe_oe_col=oe_oe_col,
                            cols_to_fill=['Вес', 'Длина', 'Ширина', 'Высота']
                        )
                        
                        if df_cross is not None and not df_cross.empty:
                            df_result = processor.build_cross_references_column(
                                df_main=df_result,
                                df_cross=df_cross,
                                main_art_col=main_art_col,
                                cross_art_col=cross_art_col,
                                cross_analog_col=cross_analog_col,
                                output_col_name="Кроссы (аналоги)"
                            )
                        
                        st.session_state.processed_catalog_df = df_result
                        
                        st.success(f"✅ Обработка завершена! {len(df_result)} строк, {len(df_result.columns)} колонок.")
                        
                        st.markdown("##### 👁️ Предпросмотр результата")
                        display_cols = [main_art_col, actual_oe_col] + [c for c in ['Вес', 'Длина', 'Ширина', 'Высота', 'Цена', 'Себестоимость'] if c in df_result.columns]
                        valid_cols = [c for c in display_cols if c in df_result.columns]
                        st.dataframe(df_result[valid_cols].head(10), use_container_width=True)
                        
                    except Exception as e:
                        st.error(f"❌ Ошибка при связывании данных: {e}")
                        logger.exception("Ошибка process_and_merge")
        else:
            st.warning("⚠️ Основной файл пуст или не удалось его прочитать.")
    
    st.divider()
    
    # --- DeepSeek AI ---
    st.subheader("🤖 DeepSeek AI")
    ds_key = key_manager.get_key("deepseek")
    if not ds_key:
        st.warning("⚠️ Ключ DeepSeek не задан. Настройте его в блоке 'Управление API ключами' выше.")
    else:
        ai_mode = st.radio(
            "Выберите задачу для AI:",
            options=["Обогащение каталога", "Актуализация тарифов", "Ничего не делать"],
            horizontal=True
        )
        
        if ai_mode == "Обогащение каталога" and not st.session_state.processed_catalog_df.empty:
            df_to_enrich = st.session_state.processed_catalog_df
            name_col = next((c for c in df_to_enrich.columns if 'наименование' in c.lower() or 'name' in c.lower()), df_to_enrich.columns[0])
            
            if st.button("🚀 Запустить обогащение каталога через DeepSeek"):
                manager = DeepSeekAPIManager(api_key=ds_key)
                st.session_state.deepseek_manager = manager
                
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                results = []
                total = min(len(df_to_enrich), 20)
                
                for i in range(total):
                    row = df_to_enrich.iloc[i]
                    status_text.text(f"Обработка {i+1}/{total}: {str(row.get(name_col, ''))[:30]}...")
                    res = manager.enrich_catalog(str(row.get(name_col, '')), str(row.get('Категория', '')))
                    results.append(res)
                    progress_bar.progress((i + 1) / total)
                    time.sleep(0.3)
                    
                status_text.text("✅ Обогащение завершено!")
                
                # Отображаем результаты
                df_enriched = df_to_enrich.head(total).copy()
                df_enriched['AI_Категория'] = [r.get('group_category', '') for r in results]
                df_enriched['AI_Тип'] = [r.get('product_type', '') for r in results]
                df_enriched['AI_Опасный'] = [r.get('hazardous', False) for r in results]
                df_enriched['AI_Хрупкий'] = [r.get('fragile', False) for r in results]
                
                st.dataframe(df_enriched[['Артикул', 'Наименование', 'AI_Категория', 'AI_Тип', 'AI_Опасный', 'AI_Хрупкий']], use_container_width=True)
                st.session_state.processed_catalog_df = df_enriched
                
        elif ai_mode == "Актуализация тарифов":
            marketplace = st.selectbox("Маркетплейс", ["Ozon", "Wildberries", "Яндекс Маркет"])
            if st.button("🚀 Получить актуальные тарифы через DeepSeek"):
                manager = DeepSeekAPIManager(api_key=ds_key)
                st.session_state.deepseek_manager = manager
                
                with st.spinner("Запрос тарифов к DeepSeek..."):
                    result = manager.update_tariffs(marketplace, "auto_parts")
                    if "error" in result:
                        st.error(f"❌ Ошибка: {result['error']}")
                    else:
                        st.success("✅ Тарифы получены!")
                        st.json(result)
                        
                        # Обновляем конфигурацию калькулятора
                        if 'fbs_calculator' in st.session_state:
                            calc = st.session_state.fbs_calculator
                            calc.config.update({
                                "commission_rate": result.get("commission_rate", calc.config.get("commission_rate", 0.15)),
                                "logistics_base": result.get("logistics_base", calc.config.get("logistics_base", 50.0)),
                                "logistics_per_kg": result.get("logistics_per_kg", calc.config.get("logistics_per_kg", 15.0)),
                                "storage_per_day": result.get("storage_per_day", calc.config.get("storage_per_day", 0.3)),
                                "acquiring_fee": result.get("acquiring_fee", calc.config.get("acquiring_fee", 0.015)),
                                "return_fee": result.get("return_fee", calc.config.get("return_fee", 0.02)),
                            })
                            st.info("✅ Тарифы обновлены в калькуляторе!")

def show_section4_calculation():
    """Раздел 4: Расчёт юнит-экономики (FBS)"""
    st.header("🧮 Раздел 4: Расчёт юнит-экономики (FBS)")
    
    st.info("""
    **В этом разделе:**
    1. Рассчитайте юнит-экономику для всех товаров в режиме FBS
    2. Экспортируйте результаты в Excel с живыми формулами
    3. Интеграция с Google Sheets для обновления остатков
    """)
    
    # --- Проверка наличия данных ---
    if st.session_state.processed_catalog_df.empty:
        st.error("❌ Нет данных каталога. Перейдите в Раздел 1 и загрузите файлы.")
        return
        
    df_catalog = st.session_state.processed_catalog_df.copy()
    
    # --- Настройки расчёта ---
    st.subheader("⚙️ Настройки расчёта")
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        marketplace = st.selectbox(
            "Маркетплейс",
            options=["Ozon", "Wildberries", "Яндекс Маркет"],
            index=0
        )
    with col2:
        operation_mode = st.selectbox(
            "Режим работы",
            options=["FBS", "FBY"],
            index=0,
            help="FBS — со своего склада по заказу. FBY — аналог FBS с доставкой МП."
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
    st.subheader("🔍 Определение колонок")
    
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
        name_col = st.selectbox("Наименование (опц.)", ["Не выбрано"] + df_catalog.columns.tolist(), index=0)
    
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
    
    calc = st.session_state.fbs_calculator
    calc.tax_system = tax_system
    
    if st.button("🚀 Рассчитать юнит-экономику для всех товаров", type="primary"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        results = []
        total = len(df_catalog)
        
        for i, (_, row) in enumerate(df_catalog.iterrows()):
            status_text.text(f"Расчёт {i+1}/{total}: {str(row.get(artikul_col, ''))[:20]}...")
            
            try:
                price = float(row.get(price_col, 0) or 0)
                cost = float(row.get(cost_col, 0) or 0)
                
                if price <= 0 or cost <= 0:
                    continue
                    
                weight = float(row.get(weight_col, 1.0) or 1.0) if weight_col != "Не выбрано" else 1.0
                length = float(row.get(length_col, 0) or 0) if length_col != "Не выбрано" else 0
                width = float(row.get(width_col, 0) or 0) if width_col != "Не выбрано" else 0
                height = float(row.get(height_col, 0) or 0) if height_col != "Не выбрано" else 0
                
                res = calc.calculate(
                    price=price,
                    cost=cost,
                    weight_kg=weight,
                    length_cm=length,
                    width_cm=width,
                    height_cm=height,
                    days_in_storage=days_in_storage,
                    operation_mode=operation_mode
                )
                
                res["Артикул"] = row.get(artikul_col, "")
                res["Наименование"] = row.get(name_col, "") if name_col != "Не выбрано" else ""
                res["weight"] = weight
                res["length"] = length
                res["width"] = width
                res["height"] = height
                res["storage_days"] = days_in_storage
                res["logistics_base"] = calc.config.get("logistics_base", 50.0)
                res["logistics_per_kg"] = calc.config.get("logistics_per_kg", 15.0)
                res["storage_rate"] = calc.config.get("storage_per_day", 0.3)
                
                results.append(res)
                
            except Exception as e:
                logger.warning(f"Ошибка расчёта для {row.get(artikul_col, '')}: {e}")
                continue
                
            progress_bar.progress((i + 1) / total)
            
        status_text.text(f"✅ Расчёт завершён! Обработано {len(results)} товаров.")
        
        if results:
            df_results = pd.DataFrame(results)
            st.session_state.calculation_results_df = df_results
            st.success(f"✅ Рассчитано {len(df_results)} товаров. Средняя маржа: {df_results['margin_percent'].mean():.1f}%")
        else:
            st.error("❌ Не удалось рассчитать ни одного товара.")
    
    # --- Отображение результатов ---
    if not st.session_state.calculation_results_df.empty:
        df_results = st.session_state.calculation_results_df
        
        st.divider()
        st.subheader("📊 Результаты расчёта")
        
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
            underpriced = (df_results['price'] < df_results['recommended_min_price']).sum()
            st.metric("💡 Недооценённых", f"{underpriced}")
        
        # Таблица результатов
        display_cols = ["Артикул", "Наименование", "price", "cost", "profit", "margin_percent", "roi", "recommended_min_price"]
        available_cols = [c for c in display_cols if c in df_results.columns]
        st.dataframe(df_results[available_cols].head(100), use_container_width=True)
        
        # --- Экспорт в Excel с живыми формулами ---
        st.divider()
        st.subheader("📥 Экспорт в Excel с живыми формулами")
        st.info("""
        **💡 КАК ЭТО РАБОТАЕТ:**
        - Жёлтые ячейки — вводные данные (можно менять)
        - Зелёные ячейки — расчётные (пересчитываются автоматически)
        - Синие ячейки — итоговые (прибыль, маржа, рек. цена)
        
        Меняйте цену, вес или ставку комиссии — и вся экономика пересчитается!
        """)
        
        if st.button("📥 Экспортировать в Excel с живыми формулами", type="primary"):
            try:
                output_path = EXPORTS_DIR / f"unit_economics_live_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                exporter = LiveExcelExporter()
                success = exporter.export_with_formulas(df_results, str(output_path), marketplace)
                
                if success and output_path.exists():
                    with open(output_path, "rb") as f:
                        file_data = f.read()
                    st.download_button(
                        label="⬇️ Скачать Excel с живыми формулами",
                        data=file_data,
                        file_name=output_path.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                    st.success("✅ Excel с живыми формулами готов к скачиванию!")
                else:
                    st.error("❌ Ошибка создания файла")
            except Exception as e:
                st.error(f"❌ Ошибка экспорта: {e}")
                logger.exception("Ошибка экспорта")
        
        # --- Google Sheets интеграция ---
        st.divider()
        st.subheader("🔄 Интеграция с Google Sheets")
        
        key_manager = st.session_state.secure_key_manager
        gs_creds = key_manager.get_key("google_sheets")
        
        if not gs_creds:
            st.warning("⚠️ Google Sheets credentials не заданы. Настройте их в Разделе 1.")
        else:
            col1, col2 = st.columns(2)
            with col1:
                spreadsheet_id = st.text_input("ID Google Таблицы", placeholder="1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms")
            with col2:
                worksheet_name = st.text_input("Название листа", value="Юнит-экономика")
            
            if st.button("📤 Экспортировать результаты в Google Sheets"):
                if not spreadsheet_id:
                    st.error("❌ Укажите ID Google Таблицы")
                else:
                    try:
                        manager = GoogleSheetsManager(gs_creds)
                        success = manager.write_sheet(
                            spreadsheet_id=spreadsheet_id,
                            df=df_results[["Артикул", "Наименование", "price", "profit", "margin_percent", "roi", "recommended_min_price"]],
                            worksheet_name=worksheet_name,
                            clear_before=True
                        )
                        if success:
                            st.success("✅ Результаты экспортированы в Google Sheets!")
                        else:
                            st.error("❌ Ошибка экспорта")
                    except Exception as e:
                        st.error(f"❌ Ошибка: {e}")
            
            # --- Обновление остатков ---
            st.markdown("---")
            st.subheader("📦 Обновление остатков через Google Sheets")
            st.info("""
            **Сценарий:**
            1. Отредактируйте остатки/цены в Google Sheets
            2. Нажмите "Загрузить изменения в МП" — система отправит данные через API
            """)
            
            ozon_api_key = key_manager.get_key("ozon")
            if not ozon_api_key:
                st.warning("⚠️ Ozon API ключ не задан. Настройте его в Разделе 1.")
            else:
                col1, col2 = st.columns(2)
                with col1:
                    ozon_client_id = st.text_input("Ozon Client ID", placeholder="12345")
                with col2:
                    stock_worksheet = st.text_input("Лист с остатками", value="Остатки")
                
                if st.button("📥 Загрузить изменения из Google Sheets в Ozon"):
                    if not spreadsheet_id or not ozon_client_id:
                        st.error("❌ Укажите ID таблицы и Client ID")
                    else:
                        try:
                            manager = GoogleSheetsManager(gs_creds)
                            df_changes = manager.read_sheet(spreadsheet_id, stock_worksheet)
                            
                            if df_changes is not None and not df_changes.empty:
                                st.write("📋 Изменения из Google Sheets:")
                                st.dataframe(df_changes.head(20))
                                
                                # Отправка в Ozon (упрощенная)
                                url = "https://api-seller.ozon.ru/v2/products/info/stocks"
                                headers = {
                                    "Client-Id": ozon_client_id,
                                    "Api-Key": ozon_api_key,
                                    "Content-Type": "application/json"
                                }
                                
                                items = []
                                for _, row in df_changes.iterrows():
                                    artikul = str(row.get("Артикул", "")).strip()
                                    stock = int(row.get("Остаток", 0))
                                    if artikul and stock >= 0:
                                        items.append({"offer_id": artikul, "stock": stock, "warehouse_id": 0})
                                
                                if items:
                                    with st.spinner(f"Отправка {len(items)} SKU в Ozon..."):
                                        results = {"updated": 0, "errors": []}
                                        for i in range(0, len(items), 100):
                                            batch = items[i:i+100]
                                            try:
                                                response = requests.post(url, json={"stocks": batch}, headers=headers, timeout=30)
                                                if response.status_code == 200:
                                                    results["updated"] += len(batch)
                                                else:
                                                    results["errors"].append(f"Batch {i//100+1}: HTTP {response.status_code}")
                                            except Exception as e:
                                                results["errors"].append(f"Batch {i//100+1}: {str(e)}")
                                        
                                        if results["updated"] > 0:
                                            st.success(f"✅ Обновлено {results['updated']} SKU!")
                                        if results["errors"]:
                                            for err in results["errors"][:5]:
                                                st.warning(f"  - {err}")
                                else:
                                    st.warning("⚠️ Нет данных для отправки")
                            else:
                                st.warning("⚠️ Нет данных для отправки")
                        except Exception as e:
                            st.error(f"❌ Ошибка: {e}")

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
    <div style='text-align: center; padding: 20px; background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%); border-radius: 10px; margin-bottom: 20px;'>
    <h1 style='color: white; margin: 0;'>{APP_NAME}</h1>
    <p style='color: #ccc; margin: 10px 0 0 0; font-size: 1.1em;'>
    {APP_DESCRIPTION}
    </p>
    <p style='color: #888; margin: 5px 0 0 0; font-size: 0.9em;'>
    Версия {APP_VERSION} | FBS-ONLY
    </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Инициализация
    init_session_state()
    
    # Sidebar навигация
    st.sidebar.title("🧭 Навигация")
    
    section = st.sidebar.radio(
        "Выберите раздел:",
        [
            "📁 Раздел 1: Загрузка и связывание данных",
            "🧮 Раздел 4: Расчёт юнит-экономики (FBS)",
        ]
    )
    
    # Информация в sidebar
    st.sidebar.markdown("---")
    st.sidebar.markdown("### 📊 Статус системы")
    
    st.sidebar.success("✅ SecureKeyManager")
    st.sidebar.success("✅ OECrossLinker")
    st.sidebar.success("✅ FBSUnitEconomicsCalculator")
    st.sidebar.success("✅ LiveExcelExporter")
    
    if CRYPTO_AVAILABLE:
        st.sidebar.success("✅ Cryptography")
    else:
        st.sidebar.warning("⚠️ Cryptography")
    
    if GSPREAD_AVAILABLE:
        st.sidebar.success("✅ GSpread")
    else:
        st.sidebar.warning("⚠️ GSpread")
    
    st.sidebar.markdown("---")
    st.sidebar.info("""
    **Быстрый старт:**
    1. 📁 Загрузите файлы (Раздел 1)
    2. 🔑 Настройте API ключи (Раздел 1)
    3. 🧮 Рассчитайте юнит-экономику (Раздел 4)
    4. 📥 Экспортируйте в Excel с живыми формулами
    """)
    
    # Отображение выбранного раздела
    if section == "📁 Раздел 1: Загрузка и связывание данных":
        show_section1_data_loading()
    elif section == "🧮 Раздел 4: Расчёт юнит-экономики (FBS)":
        show_section4_calculation()
    
    # Футер
    st.divider()
    st.markdown(f"""
    <div style='text-align: center; padding: 20px; color: #666;'>
    <p style='margin: 0;'>🚀 <strong>FBS Юнит-экономика PRO 2026</strong></p>
    <p style='margin: 5px 0 0 0; font-size: 0.9em;'>
    Версия {APP_VERSION} | Живые формулы Excel | Безопасное хранение ключей
    </p>
    </div>
    """, unsafe_allow_html=True)

if __name__ == "__main__":
    main()
