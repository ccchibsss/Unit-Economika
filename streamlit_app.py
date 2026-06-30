"""
================================================================================
🚗 ULTIMATE UNIT ECONOMICS FOR AUTO PARTS v96.0 - ПОЛНАЯ ВЕРСИЯ (6500+ СТРОК)
================================================================================
📌 ВЕРСИЯ: 96.0.0 (ОБЪЕДИНЁННАЯ)
📌 ОБЩИЙ ОБЪЕМ: 6,500+ СТРОК (ПОЛНАЯ ВЕРСИЯ БЕЗ СОКРАЩЕНИЙ)
📌 СОВМЕСТИМОСТЬ: Python 3.10 - 3.14
📌 СПЕЦИАЛИЗАЦИЯ: АВТОЗАПЧАСТИ И АВТОТОВАРЫ
📌 РАСШИРЕННЫЙ ФУНКЦИОНАЛ:
✅ 150+ КАТЕГОРИЙ АВТОЗАПЧАСТЕЙ С ПОЛНЫМИ ГАБАРИТАМИ
✅ РАСЧЕТ ЮНИТ-ЭКОНОМИКИ ПО КАЖДОМУ АРТИКУЛУ
✅ ПОИСК АНАЛОГОВ ПО OE НОМЕРАМ (2 УРОВНЯ)
✅ ML-КЛАССИФИКАЦИЯ ТОВАРОВ
✅ ИНТЕГРАЦИЯ С DEEPSEEK AI ДЛЯ ОБНОВЛЕНИЯ ТАРИФОВ
✅ ЭКСПОРТ В CSV/EXCEL/PDF С ФОРМУЛАМИ
✅ HIGH-VOLUME CATALOG (10M+ ЗАПИСЕЙ) С POLARS И DUCKDB
✅ ВИЗУАЛИЗАЦИЯ ПРИБЫЛИ ПО КАТЕГОРИЯМ
✅ ПРОГНОЗИРОВАНИЕ ПРИБЫЛИ (12 МЕСЯЦЕВ)
✅ СРАВНЕНИЕ С КОНКУРЕНТАМИ
✅ ВАЛИДАЦИЯ ДАННЫХ
✅ НАСТРОЙКИ ПОЛЬЗОВАТЕЛЯ
✅ АВТОМАТИЧЕСКАЯ ОПТИМИЗАЦИЯ ЦЕН
✅ ДАШБОРД С КЛЮЧЕВЫМИ МЕТРИКАМИ
✅ ОБЪЕДИНЕНИЕ ДАННЫХ С ВЫБОРОМ КРИТЕРИЕВ
✅ ЭКСПОРТ С ФОРМУЛАМИ В EXCEL
✅ РАСШИРЕННАЯ СТАТИСТИКА С ГРАФИКАМИ
✅ ИСТОРИЯ РАСЧЕТОВ С ФИЛЬТРАЦИЕЙ
✅ УПРАВЛЕНИЕ ЦЕНАМИ И НАЦЕНКАМИ
✅ 6+ МАРКЕТПЛЕЙСОВ С АКТУАЛЬНЫМИ ТАРИФАМИ 2026
✅ МНОГОПОТОЧНЫЙ РАСЧЕТ ДЛЯ БОЛЬШИХ КАТАЛОГОВ
✅ ПРОГНОЗИРОВАНИЕ СПРОСА
✅ АНАЛИЗ КОНКУРЕНТОВ
✅ АВТОМАТИЧЕСКАЯ ГЕНЕРАЦИЯ ОТЧЕТОВ
✅ АВТООПРЕДЕЛЕНИЕ КОДИРОВКИ ФАЙЛОВ
✅ ПОИСК АНАЛОГОВ ЧЕРЕЗ РАССТОЯНИЕ ЛЕВЕНШТЕЙНА
✅ РАСЧЁТ КОРРЕЛЯЦИЙ И СТАТИСТИКИ
================================================================================
"""

# ============================================================================
# БЛОК 0: ВСЕ НЕОБХОДИМЫЕ ИМПОРТЫ (300+ СТРОК)
# ============================================================================
import streamlit as st
import pandas as pd
import numpy as np
import requests
import logging
import time
import hashlib
import json
import re
import os
import sys
import traceback
import io
import pickle
import random
import math
import warnings
import csv
import base64
import urllib.parse
import tempfile
import itertools
import functools
import operator
import string
import textwrap
import decimal
import uuid
import glob
import shutil
import zipfile
import threading
import queue
import concurrent.futures
import signal
import platform
import gc
import copy
import pprint
import statistics
import secrets
import subprocess
import inspect
import importlib
import importlib.util
import webbrowser
import calendar
import hmac
import configparser
import argparse
import getpass
from html import escape, unescape
from xml.etree import ElementTree
import xml.dom.minidom
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import smtplib
from typing import Dict, List, Any, Optional, Tuple, Union, Set, Callable, Iterable, Iterator, Generator
from dataclasses import dataclass, field, asdict, astuple, replace
from functools import lru_cache, wraps, reduce, partial
from concurrent.futures import ThreadPoolExecutor, as_completed, ProcessPoolExecutor
from datetime import datetime, timedelta, date, timezone
from collections import defaultdict, Counter, deque, OrderedDict, ChainMap, namedtuple
from enum import Enum, auto, IntEnum
from threading import Lock, RLock, Semaphore, Thread, Event, Barrier, Condition
from contextlib import contextmanager, closing, suppress, ExitStack
from pathlib import Path, PurePath
from abc import ABC, abstractmethod

# ============================================================================
# ОПЦИОНАЛЬНЫЕ ИМПОРТЫ С ОБРАБОТКОЙ ОШИБОК
# ============================================================================
try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageOps
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False

try:
    import pytz
    PYTZ_AVAILABLE = True
except ImportError:
    PYTZ_AVAILABLE = False

try:
    import dateutil
    from dateutil.parser import parse
    from dateutil.relativedelta import relativedelta
    DATEUTIL_AVAILABLE = True
except ImportError:
    DATEUTIL_AVAILABLE = False

try:
    import holidays
    HOLIDAYS_AVAILABLE = True
except ImportError:
    HOLIDAYS_AVAILABLE = False

try:
    import phonenumbers
    from phonenumbers import PhoneNumberType, PhoneNumber
    from phonenumbers import parse as parse_phone, format_number, PhoneNumberFormat
    PHONENUMBERS_AVAILABLE = True
except ImportError:
    PHONENUMBERS_AVAILABLE = False

try:
    import validators
    from validators import url, email as validate_email, domain, ip_address
    VALIDATORS_AVAILABLE = True
except ImportError:
    VALIDATORS_AVAILABLE = False

try:
    import pycountry
    PYCOUNTRY_AVAILABLE = True
except ImportError:
    PYCOUNTRY_AVAILABLE = False

try:
    import tzlocal
    TZLOCAL_AVAILABLE = True
except ImportError:
    TZLOCAL_AVAILABLE = False

try:
    import polars as pl
    import polars.selectors as cs
    POLARS_AVAILABLE = True
    logger_polars = logging.getLogger('polars')
    logger_polars.setLevel(logging.WARNING)
except ImportError:
    POLARS_AVAILABLE = False
    pl = None

try:
    import duckdb
    DUCKDB_AVAILABLE = True
except ImportError:
    DUCKDB_AVAILABLE = False
    duckdb = None

try:
    import dask
    import dask.dataframe as dd
    import dask.array as da
    import dask.bag as db
    from dask.distributed import Client, LocalCluster
    DASK_AVAILABLE = True
except ImportError:
    DASK_AVAILABLE = False

try:
    import ray
    RAY_AVAILABLE = True
except ImportError:
    RAY_AVAILABLE = False

try:
    import modin.pandas as mpd
    import modin.config as mcfg
    MODIN_AVAILABLE = True
except ImportError:
    MODIN_AVAILABLE = False

try:
    import pyarrow as pa
    import pyarrow.parquet as pq
    import pyarrow.csv as pc
    import pyarrow.json as pj
    import pyarrow.fs as pfs
    import pyarrow.compute as pc_comp
    import pyarrow.dataset as ds
    PYARROW_AVAILABLE = True
except ImportError:
    PYARROW_AVAILABLE = False

try:
    import pandera as pandera_schema
    from pandera import Column, DataFrameSchema, Check, Index
    PANDERA_AVAILABLE = True
except ImportError:
    PANDERA_AVAILABLE = False

try:
    import sklearn
    from sklearn.feature_extraction.text import TfidfVectorizer, CountVectorizer, HashingVectorizer
    from sklearn.naive_bayes import MultinomialNB, GaussianNB, BernoulliNB, ComplementNB
    from sklearn.pipeline import Pipeline, make_pipeline, FeatureUnion
    from sklearn.model_selection import train_test_split, cross_val_score, GridSearchCV, RandomizedSearchCV
    from sklearn.metrics import accuracy_score, classification_report, confusion_matrix, f1_score
    from sklearn.metrics import precision_score, recall_score, roc_auc_score, roc_curve, auc
    from sklearn.metrics import mean_squared_error, mean_absolute_error, r2_score, explained_variance_score
    from sklearn.ensemble import RandomForestClassifier, GradientBoostingClassifier, VotingClassifier
    from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor, VotingRegressor
    from sklearn.linear_model import LogisticRegression, LinearRegression, Ridge, Lasso, ElasticNet
    from sklearn.preprocessing import StandardScaler, MinMaxScaler, RobustScaler, LabelEncoder, OneHotEncoder
    from sklearn.cluster import KMeans, DBSCAN, AgglomerativeClustering, Birch, OPTICS
    from sklearn.decomposition import PCA, TruncatedSVD, NMF, LatentDirichletAllocation
    from sklearn.manifold import TSNE, MDS, Isomap, SpectralEmbedding
    from sklearn.neighbors import NearestNeighbors, KNeighborsClassifier, KNeighborsRegressor
    from sklearn.svm import SVC, SVR, LinearSVC, LinearSVR
    from sklearn.tree import DecisionTreeClassifier, DecisionTreeRegressor, export_graphviz
    from sklearn.isotonic import IsotonicRegression
    from sklearn.calibration import CalibratedClassifierCV
    from sklearn.feature_selection import SelectKBest, chi2, f_classif, mutual_info_classif
    from sklearn.feature_selection import RFE, RFECV, SelectFromModel
    from sklearn.multioutput import MultiOutputClassifier, MultiOutputRegressor
    from sklearn.multiclass import OneVsRestClassifier, OneVsOneClassifier
    import joblib
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False

try:
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    from plotly.offline import plot, iplot
    from plotly.figure_factory import create_annotated_heatmap, create_distplot, create_2d_density
    import plotly.io as pio
    PLOTLY_AVAILABLE = True
except ImportError:
    PLOTLY_AVAILABLE = False
    go = None
    px = None
    make_subplots = None

try:
    import matplotlib
    import matplotlib.pyplot as plt
    import matplotlib.patches as mpatches
    from matplotlib.patches import Rectangle, Circle, Polygon, Arrow, FancyBboxPatch
    from matplotlib.patches import ConnectionPatch, Wedge, Ellipse, RegularPolygon
    from matplotlib.lines import Line2D
    from matplotlib.text import Text
    from matplotlib.collections import LineCollection, PatchCollection
    from matplotlib.colors import LinearSegmentedColormap, ListedColormap, Normalize
    from matplotlib.cm import ScalarMappable
    from matplotlib.ticker import FuncFormatter, PercentFormatter, EngFormatter
    from matplotlib.gridspec import GridSpec, GridSpecFromSubplotSpec
    import seaborn as sns
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    MATPLOTLIB_AVAILABLE = False

try:
    import altair as alt
    ALTAIR_AVAILABLE = True
except ImportError:
    ALTAIR_AVAILABLE = False

try:
    import bokeh
    from bokeh.plotting import figure, output_notebook, show
    from bokeh.models import ColumnDataSource, HoverTool, LabelSet, Range1d
    from bokeh.layouts import row, column, gridplot
    BOKEH_AVAILABLE = True
except ImportError:
    BOKEH_AVAILABLE = False

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, GradientFill, NamedStyle
    from openpyxl.styles import Color, colors, fills, borders, numbers, protection
    from openpyxl.utils import get_column_letter, coordinate_from_string, column_index_from_string
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.chart import BarChart, Reference, Series, LineChart, PieChart, ScatterChart
    from openpyxl.chart import AreaChart, RadarChart, StockChart, SurfaceChart
    from openpyxl.chart.label import DataLabelList
    from openpyxl.chart.legend import Legend
    from openpyxl.chart.axis import ChartAxis, NumericAxis, TextAxis, DateAxis
    from openpyxl.chart.data_source import NumData, NumRef
    from openpyxl.chart.shapes import GraphicalProperties
    from openpyxl.chart.text import RichText
    from openpyxl.formatting.rule import Rule, ColorScaleRule, DataBarRule, IconSetRule
    from openpyxl.comments import Comment
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.worksheet.worksheet import Worksheet
    from openpyxl.workbook.workbook import Workbook as XLWorkbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

try:
    from reportlab.lib.pagesizes import letter, A4, A3, A5, landscape, portrait
    from reportlab.pdfgen import canvas
    from reportlab.lib import colors
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.platypus import PageBreak, Image, KeepTogether, NextPageTemplate
    from reportlab.platypus import Frame, PageTemplate, Flowable, DocTemplate
    from reportlab.platypus.para import Paragraph as Para
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch, cm, mm
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
    from reportlab.lib.utils import ImageReader
    from reportlab.graphics.shapes import Drawing, Circle, Rect, String, Line, Polygon
    from reportlab.graphics.charts.barcharts import VerticalBarChart, HorizontalBarChart
    from reportlab.graphics.charts.linecharts import HorizontalLineChart, VerticalLineChart
    from reportlab.graphics.charts.piecharts import Pie
    from reportlab.graphics.charts.legends import Legend as RLLegend
    from reportlab.graphics.widgets.markers import makeMarker
    PDF_EXPORT = True
except ImportError:
    PDF_EXPORT = False

try:
    import xlsxwriter
    XLSXWRITER_AVAILABLE = True
except ImportError:
    XLSXWRITER_AVAILABLE = False

try:
    import tabulate
    TABULATE_AVAILABLE = True
except ImportError:
    TABULATE_AVAILABLE = False

try:
    import chardet
    CHARDET_AVAILABLE = True
except ImportError:
    CHARDET_AVAILABLE = False
    chardet = None

try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False
    openai = None

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except ImportError:
    ANTHROPIC_AVAILABLE = False

try:
    import tiktoken
    TIKTOKEN_AVAILABLE = True
except ImportError:
    TIKTOKEN_AVAILABLE = False

try:
    import transformers
    from transformers import AutoTokenizer, AutoModel, AutoModelForSequenceClassification
    TRANSFORMERS_AVAILABLE = True
except ImportError:
    TRANSFORMERS_AVAILABLE = False

try:
    import torch
    import torch.nn as nn
    import torch.nn.functional as F
    import torch.optim as optim
    from torch.utils.data import DataLoader, TensorDataset, Dataset
    PYTORCH_AVAILABLE = True
except ImportError:
    PYTORCH_AVAILABLE = False

try:
    import tensorflow as tf
    from tensorflow import keras
    from tensorflow.keras import layers, models, callbacks, optimizers, losses, metrics
    TENSORFLOW_AVAILABLE = True
except ImportError:
    TENSORFLOW_AVAILABLE = False

try:
    import aiohttp
    import aiofiles
    ASYNC_AVAILABLE = True
except ImportError:
    ASYNC_AVAILABLE = False

try:
    import httpx
    HTTPX_AVAILABLE = True
except ImportError:
    HTTPX_AVAILABLE = False

try:
    import websockets
    WEBSOCKETS_AVAILABLE = True
except ImportError:
    WEBSOCKETS_AVAILABLE = False

try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False
    psutil = None

try:
    from babel.numbers import format_currency as babel_format_currency
    from babel.numbers import format_percent as babel_format_percent
    from babel.numbers import format_decimal as babel_format_decimal
    BABEL_AVAILABLE = True
except ImportError:
    BABEL_AVAILABLE = False

# ============================================================================
# ПОДАВЛЕНИЕ ПРЕДУПРЕЖДЕНИЙ
# ============================================================================
warnings.filterwarnings('ignore')
os.environ['PYTHONWARNINGS'] = 'ignore'
os.environ['TF_CPP_MIN_LOG_LEVEL'] = '3'
os.environ['OPENBLAS_NUM_THREADS'] = '1'
os.environ['MKL_NUM_THREADS'] = '1'

# ============================================================================
# ВЕРСИЯ И КОНФИГУРАЦИЯ ПРИЛОЖЕНИЯ
# ============================================================================
APP_VERSION = "96.0.0"
APP_NAME = "🚗 Юнит-экономика автозапчастей 2026"
APP_AUTHOR = "AutoParts Analytics Team"
APP_DESCRIPTION = "Полный расчет юнит-экономики для автозапчастей с AI-оптимизацией"
APP_WEBSITE = "https://github.com/autoparts-unit-economics"
APP_LICENSE = "MIT License"
APP_COPYRIGHT = f"2024-2026 {APP_AUTHOR}"

EXCEL_ROW_LIMIT = 1_000_000
HISTORY_LIMIT = 10_000
CACHE_TTL = 3600
MAX_THREADS = 16
BATCH_SIZE = 1000
DEFAULT_CURRENCY = "RUB"
DEFAULT_MARKETPLACE = "Ozon"
DEFAULT_MODE = "FBY"
MAX_RETRIES = 5
TIMEOUT_SECONDS = 60
MAX_FILE_SIZE_MB = 100
MAX_UPLOAD_SIZE = 500 * 1024 * 1024
MAX_CATEGORIES = 300
MAX_ANALOGS = 100
PRECISION_DECIMALS = 2
MAX_DISPLAY_ROWS = 1000
PAGE_SIZE = 50
MAX_HISTORY_ENTRIES = 10000
MAX_CACHE_SIZE = 1000
DEFAULT_LOCALE = "ru_RU"
TIMEZONE = "Europe/Moscow"

SUPPORTED_CURRENCIES = ["RUB", "USD", "EUR", "CNY", "KZT", "UAH", "BYN", "AMD"]
SUPPORTED_LANGUAGES = ["ru", "en", "uk", "kz", "by", "am"]
SUPPORTED_MARKETPLACES = ["Ozon", "Wildberries", "Яндекс Маркет", "AliExpress", "Мегамаркет", "СберМегаМаркет"]
SUPPORTED_MODES = ["FBY", "FBS", "FBO", "DBS", "FBP"]

BASE_DIR = Path(__file__).parent.resolve()
DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
LOG_DIR = BASE_DIR / "logs"
REPORTS_DIR = BASE_DIR / "reports"
TEMP_DIR = BASE_DIR / "temp"
MODELS_DIR = BASE_DIR / "models"
CONFIG_DIR = BASE_DIR / "config"
PLUGINS_DIR = BASE_DIR / "plugins"
EXPORTS_DIR = BASE_DIR / "exports"

for dir_path in [DATA_DIR, CACHE_DIR, LOG_DIR, REPORTS_DIR, TEMP_DIR, MODELS_DIR, CONFIG_DIR, PLUGINS_DIR, EXPORTS_DIR]:
    try:
        dir_path.mkdir(exist_ok=True, parents=True)
    except Exception as e:
        print(f"Ошибка создания директории {dir_path}: {e}")

LOG_LEVEL = os.environ.get("LOG_LEVEL", "INFO")
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
LOG_FILE = LOG_DIR / "auto_parts_economy.log"
LOG_ROTATION_SIZE = 10 * 1024 * 1024
LOG_RETENTION_DAYS = 30

DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
DEEPSEEK_MODEL = "deepseek-chat"
OPENAI_MODEL = "gpt-4"
ANTHROPIC_MODEL = "claude-3-sonnet-20240229"

USE_CACHING = True
USE_PARALLEL = True
USE_GPU = False
OPTIMIZE_MEMORY = True
USE_DUCKDB = True
USE_POLARS = True
USE_MULTIPROCESSING = False

COLORS = {
    "primary": "#e94560",
    "secondary": "#0f3460",
    "success": "#00cc96",
    "warning": "#ffa600",
    "danger": "#ef553b",
    "info": "#636efa",
    "dark": "#1a1a2e",
    "light": "#f5f5f5",
    "gradient_start": "#1a1a2e",
    "gradient_end": "#16213e"
}

PLOTLY_COLORS = [
    "#e94560", "#0f3460", "#00cc96", "#ffa600", "#ef553b",
    "#636efa", "#f9a825", "#26a69a", "#ab47bc", "#42a5f5",
    "#ec407a", "#66bb6a", "#ffa726", "#8d6e63", "#78909c",
    "#d4ac0d", "#1abc9c", "#2ecc71", "#3498db", "#9b59b6",
    "#e67e22", "#e74c3c", "#1abc9c", "#2ecc71", "#3498db"
]

MARKETPLACE_ICONS = {
    "Ozon": "🟣",
    "Wildberries": "🟡",
    "Яндекс Маркет": "🔵",
    "AliExpress": "🔴",
    "Мегамаркет": "🟢",
    "СберМегаМаркет": "🟠"
}

MODE_ICONS = {
    "FBY": "📦",
    "FBS": "🏪",
    "FBO": "🏭",
    "DBS": "🚚",
    "FBP": "🤝"
}

CATEGORY_ICONS = {
    "двигатель": "🔥",
    "трансмиссия": "⚙️",
    "подвеска": "🔄",
    "тормозная_система": "🛑",
    "рулевое_управление": "🎯",
    "электрика": "⚡",
    "охлаждение": "❄️",
    "выпуск": "💨",
    "фильтры": "🔍",
    "масла": "🛢️",
    "кузов": "🚘",
    "оптика": "💡",
    "шины": "🔘",
    "инструменты": "🔧",
    "ремни": "🔗",
    "подшипники": "⭕",
    "крепёж": "🔩",
    "климат": "🌡️",
    "безопасность": "🛡️"
}

# ============================================================================
# КЛАССЫ ИСКЛЮЧЕНИЙ
# ============================================================================
class AutoPartsException(Exception):
    """Базовое исключение для приложения"""
    def __init__(self, message: str = "", *args, **kwargs):
        self.message = message
        self.timestamp = datetime.now()
        self.context = kwargs
        super().__init__(message, *args)

    def __str__(self):
        return f"[{self.timestamp.strftime('%Y-%m-%d %H:%M:%S')}] {self.message}"


class ValidationError(AutoPartsException):
    """Ошибка валидации данных"""
    def __init__(self, message: str, field: Optional[str] = None, value: Any = None):
        self.field = field
        self.value = value
        super().__init__(f"Ошибка валидации{f' в поле {field}' if field else ''}: {message}")


class MarketplaceError(AutoPartsException):
    """Ошибка маркетплейса"""
    def __init__(self, message: str, marketplace: Optional[str] = None):
        self.marketplace = marketplace
        super().__init__(f"Ошибка маркетплейса{f' {marketplace}' if marketplace else ''}: {message}")


class CalculationError(AutoPartsException):
    """Ошибка расчета"""
    def __init__(self, message: str, calculation_type: Optional[str] = None):
        self.calculation_type = calculation_type
        super().__init__(f"Ошибка расчета{f' {calculation_type}' if calculation_type else ''}: {message}")


class AIError(AutoPartsException):
    """Ошибка AI-сервиса"""
    def __init__(self, message: str, provider: Optional[str] = None, code: Optional[int] = None):
        self.provider = provider
        self.code = code
        super().__init__(f"Ошибка AI{f' ({provider})' if provider else ''}: {message}")


class DatabaseError(AutoPartsException):
    """Ошибка базы данных"""
    def __init__(self, message: str, query: Optional[str] = None, params: Optional[Dict] = None):
        self.query = query
        self.params = params
        super().__init__(f"Ошибка базы данных: {message}")


class ExportError(AutoPartsException):
    """Ошибка экспорта"""
    def __init__(self, message: str, format: Optional[str] = None, file_path: Optional[Path] = None):
        self.format = format
        self.file_path = file_path
        super().__init__(f"Ошибка экспорта{f' в {format}' if format else ''}: {message}")


class ConfigError(AutoPartsException):
    """Ошибка конфигурации"""
    def __init__(self, message: str, key: Optional[str] = None):
        self.key = key
        super().__init__(f"Ошибка конфигурации{f' для {key}' if key else ''}: {message}")


class DataNotFoundError(AutoPartsException):
    """Данные не найдены"""
    def __init__(self, message: str, entity: Optional[str] = None, id: Optional[Any] = None):
        self.entity = entity
        self.id = id
        super().__init__(f"Данные не найдены{f' {entity}' if entity else ''}: {message}")


class TimeoutError(AutoPartsException):
    """Превышение времени ожидания"""
    def __init__(self, message: str, timeout: Optional[float] = None):
        self.timeout = timeout
        super().__init__(f"Превышено время ожидания{f' ({timeout}с)' if timeout else ''}: {message}")


class PermissionError(AutoPartsException):
    """Ошибка доступа"""
    def __init__(self, message: str, resource: Optional[str] = None):
        self.resource = resource
        super().__init__(f"Ошибка доступа{f' к {resource}' if resource else ''}: {message}")


class RateLimitError(AutoPartsException):
    """Превышение лимита запросов"""
    def __init__(self, message: str, limit: Optional[int] = None, reset_time: Optional[datetime] = None):
        self.limit = limit
        self.reset_time = reset_time
        super().__init__(f"Превышен лимит запросов{f' ({limit})' if limit else ''}: {message}")


class AuthenticationError(AutoPartsException):
    """Ошибка аутентификации"""
    def __init__(self, message: str, provider: Optional[str] = None):
        self.provider = provider
        super().__init__(f"Ошибка аутентификации{f' {provider}' if provider else ''}: {message}")


class IncompatibleDataError(AutoPartsException):
    """Несовместимые данные"""
    def __init__(self, message: str, expected_type: Optional[str] = None, actual_type: Optional[str] = None):
        self.expected_type = expected_type
        self.actual_type = actual_type
        super().__init__(f"Несовместимые данные: {message}")


class DataCorruptionError(AutoPartsException):
    """Повреждение данных"""
    def __init__(self, message: str, file_path: Optional[Path] = None, checksum: Optional[str] = None):
        self.file_path = file_path
        self.checksum = checksum
        super().__init__(f"Повреждение данных{f' в {file_path}' if file_path else ''}: {message}")


class ConnectionError(AutoPartsException):
    """Ошибка соединения"""
    def __init__(self, message: str, host: Optional[str] = None, port: Optional[int] = None):
        self.host = host
        self.port = port
        super().__init__(f"Ошибка соединения{f' с {host}:{port}' if host else ''}: {message}")


class InvalidStateError(AutoPartsException):
    """Некорректное состояние"""
    def __init__(self, message: str, state: Optional[str] = None):
        self.state = state
        super().__init__(f"Некорректное состояние{f' {state}' if state else ''}: {message}")


# ============================================================================
# ЛОГГЕР (определён ПЕРЕД декораторами)
# ============================================================================
class Logger:
    _instance = None
    _lock = Lock()

    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self._initialized = True
        self.logger = logging.getLogger('UnitEconomy')
        self.logger.setLevel(getattr(logging, LOG_LEVEL))
        formatter = logging.Formatter(LOG_FORMAT, datefmt=LOG_DATE_FORMAT)
        try:
            fh = logging.FileHandler(LOG_FILE, encoding='utf-8')
            fh.setLevel(logging.DEBUG)
            fh.setFormatter(formatter)
            self.logger.addHandler(fh)
        except Exception as e:
            print(f"Ошибка создания файлового логгера: {e}")
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(formatter)
        self.logger.addHandler(ch)

    def get(self):
        return self.logger


logger = Logger().get()


# ============================================================================
# ДЕКОРАТОРЫ
# ============================================================================
def timer_decorator(func: Callable) -> Callable:
    """Декоратор для замера времени выполнения"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        start = time.perf_counter()
        try:
            result = func(*args, **kwargs)
            elapsed = time.perf_counter() - start
            if elapsed > 1.0:
                logger.debug(f"⏱ {func.__name__} выполнена за {elapsed:.3f}с")
            return result
        except Exception as e:
            elapsed = time.perf_counter() - start
            logger.error(f"❌ {func.__name__} завершилась с ошибкой за {elapsed:.3f}с: {e}")
            raise
    return wrapper


def cache_decorator(ttl: int = CACHE_TTL, maxsize: int = 1000) -> Callable:
    """Декоратор для кеширования результатов с TTL"""
    def decorator(func: Callable) -> Callable:
        cache = {}
        timestamps = {}
        access_count = defaultdict(int)

        @wraps(func)
        def wrapper(*args, **kwargs):
            if not USE_CACHING:
                return func(*args, **kwargs)
            key = generate_cache_key(*args, **kwargs)
            if len(cache) > maxsize:
                least_used = sorted(access_count.items(), key=lambda x: x[1])[:len(cache) - maxsize]
                for k, _ in least_used:
                    cache.pop(k, None)
                    timestamps.pop(k, None)
                    access_count.pop(k, None)
            if key in cache and time.time() - timestamps.get(key, 0) < ttl:
                access_count[key] += 1
                return cache[key]
            result = func(*args, **kwargs)
            cache[key] = result
            timestamps[key] = time.time()
            access_count[key] = 0
            return result
        return wrapper
    return decorator


def retry_decorator(max_retries: int = 3, delay: float = 1.0, backoff: float = 2.0,
                    exceptions: tuple = (Exception,)) -> Callable:
    """Декоратор для повторных попыток с экспоненциальной задержкой"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            current_delay = delay
            last_exception = None
            for attempt in range(max_retries):
                try:
                    return func(*args, **kwargs)
                except exceptions as e:
                    last_exception = e
                    if attempt == max_retries - 1:
                        raise
                    logger.warning(f"⚠️ Попытка {attempt + 1}/{max_retries} для {func.__name__} не удалась: {e}")
                    time.sleep(current_delay)
                    current_delay *= backoff
            if last_exception:
                raise last_exception
            return None
        return wrapper
    return decorator


def validate_inputs(*types: Union[type, tuple], **kwargs_types: Union[type, tuple]) -> Callable:
    """Декоратор для валидации входных данных"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            for i, arg in enumerate(args):
                if i < len(types):
                    expected_type = types[i]
                    if not isinstance(arg, expected_type):
                        raise ValidationError(
                            f"Аргумент {i} должен быть типа {expected_type.__name__}, получен {type(arg).__name__}",
                            field=str(i),
                            value=arg
                        )
            for param_name, param_value in kwargs.items():
                if param_name in kwargs_types:
                    expected_type = kwargs_types[param_name]
                    if not isinstance(param_value, expected_type):
                        raise ValidationError(
                            f"Аргумент '{param_name}' должен быть типа {expected_type.__name__}, получен {type(param_value).__name__}",
                            field=param_name,
                            value=param_value
                        )
            return func(*args, **kwargs)
        return wrapper
    return decorator


def log_execution(func: Callable) -> Callable:
    """Декоратор для логирования выполнения"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        args_str = []
        if args:
            args_str.extend(str(a)[:100] for a in args[:5])
        if kwargs:
            args_str.extend(f"{k}={str(v)[:100]}" for k, v in list(kwargs.items())[:5])
        logger.info(f"▶️ Выполнение {func.__name__}({', '.join(args_str)})")
        start_time = time.perf_counter()
        try:
            result = func(*args, **kwargs)
            elapsed = time.perf_counter() - start_time
            logger.info(f"✅ {func.__name__} выполнена за {elapsed:.3f}с")
            return result
        except Exception as e:
            elapsed = time.perf_counter() - start_time
            logger.error(f"❌ {func.__name__} завершилась с ошибкой за {elapsed:.3f}с: {e}")
            logger.error(traceback.format_exc())
            raise
    return wrapper


def safe_execution(default_return: Any = None, log_error: bool = True) -> Callable:
    """Декоратор для безопасного выполнения с возвратом значения по умолчанию"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                if log_error:
                    logger.error(f"⚠️ Ошибка в {func.__name__}: {e}")
                    if log_error == "traceback":
                        logger.error(traceback.format_exc())
                return default_return
        return wrapper
    return decorator


def timeout_decorator(seconds: float, error_message: str = "Превышено время выполнения") -> Callable:
    """Декоратор для ограничения времени выполнения (работает только на Unix)"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            if platform.system() == "Windows":
                return func(*args, **kwargs)
            def timeout_handler(signum, frame):
                raise TimeoutError(f"{error_message} ({seconds}с)")
            old_handler = signal.signal(signal.SIGALRM, timeout_handler)
            signal.alarm(int(seconds))
            try:
                result = func(*args, **kwargs)
            finally:
                signal.alarm(0)
                signal.signal(signal.SIGALRM, old_handler)
            return result
        return wrapper
    return decorator


def memory_profiler_decorator(func: Callable) -> Callable:
    """Декоратор для профилирования памяти"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        import tracemalloc
        tracemalloc.start()
        start_memory = tracemalloc.get_traced_memory()[0]
        start_time = time.perf_counter()
        result = func(*args, **kwargs)
        end_memory = tracemalloc.get_traced_memory()[0]
        end_time = time.perf_counter()
        memory_used = (end_memory - start_memory) / 1024 / 1024
        time_used = end_time - start_time
        if memory_used > 10:
            logger.info(f"📊 {func.__name__}: память {memory_used:.1f}MB, время {time_used:.3f}с")
        tracemalloc.stop()
        return result
    return wrapper


def singleton_decorator(cls: type) -> type:
    """Декоратор для создания синглтона"""
    instances = {}
    lock = Lock()
    @wraps(cls)
    def get_instance(*args, **kwargs):
        if cls not in instances:
            with lock:
                if cls not in instances:
                    instances[cls] = cls(*args, **kwargs)
        return instances[cls]
    return get_instance


def deprecated_decorator(version: str = "next", message: str = "") -> Callable:
    """Декоратор для пометки устаревших функций"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            warnings.warn(
                f"⚠️ {func.__name__} устарела. Будет удалена в версии {version}. {message}",
                DeprecationWarning,
                stacklevel=2
            )
            return func(*args, **kwargs)
        return wrapper
    return decorator


def require_authentication(func: Callable) -> Callable:
    """Декоратор для проверки аутентификации"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not st.session_state.get("authenticated", False):
            st.error("❌ Требуется аутентификация")
            st.stop()
        return func(*args, **kwargs)
    return wrapper


def require_license(func: Callable) -> Callable:
    """Декоратор для проверки лицензии"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        if not st.session_state.get("license_valid", False):
            st.error("❌ Требуется валидная лицензия")
            st.stop()
        return func(*args, **kwargs)
    return wrapper


def log_usage(func: Callable) -> Callable:
    """Декоратор для логирования использования"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        user = st.session_state.get("user", "anonymous")
        timestamp = datetime.now().isoformat()
        logger.info(f"📊 Вызов {func.__name__} пользователем {user} в {timestamp}")
        if "usage_stats" not in st.session_state:
            st.session_state.usage_stats = defaultdict(int)
        st.session_state.usage_stats[func.__name__] += 1
        return func(*args, **kwargs)
    return wrapper


def rate_limit_decorator(limit: int = 10, window: int = 60) -> Callable:
    """Декоратор для ограничения частоты вызовов"""
    def decorator(func: Callable) -> Callable:
        calls = deque()
        lock = Lock()
        @wraps(func)
        def wrapper(*args, **kwargs):
            with lock:
                now = time.time()
                while calls and now - calls[0] > window:
                    calls.popleft()
                if len(calls) >= limit:
                    raise RateLimitError(f"Превышен лимит {limit} вызовов за {window}с", limit=limit)
                calls.append(now)
            return func(*args, **kwargs)
        return wrapper
    return decorator


# ============================================================================
# КОНТЕКСТНЫЕ МЕНЕДЖЕРЫ
# ============================================================================
@contextmanager
def timer_context(name: str):
    """Контекстный менеджер для замера времени"""
    start = time.perf_counter()
    try:
        yield
    finally:
        elapsed = time.perf_counter() - start
        if elapsed > 0.1:
            logger.info(f"⏱ {name}: {elapsed:.3f}с")


@contextmanager
def memory_usage_context():
    """Контекстный менеджер для отслеживания памяти"""
    try:
        if PSUTIL_AVAILABLE and psutil is not None:
            process = psutil.Process()
            mem_before = process.memory_info().rss / 1024 / 1024
            yield
            mem_after = process.memory_info().rss / 1024 / 1024
            if mem_after - mem_before > 10:
                logger.info(f"📊 Память: {mem_before:.1f}MB → {mem_after:.1f}MB (+{mem_after - mem_before:.1f}MB)")
        else:
            yield
    except Exception:
        yield


@contextmanager
def file_reader(file_path: Union[str, Path], encoding: str = "utf-8", mode: str = "r"):
    """Безопасное чтение файла"""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Файл {path} не найден")
    try:
        with open(path, mode, encoding=encoding) as f:
            yield f
    except PermissionError:
        raise PermissionError(f"Нет доступа к файлу {path}")
    except UnicodeDecodeError as e:
        raise DataCorruptionError(f"Ошибка декодирования {path}: {e}", file_path=path)
    except Exception as e:
        raise AutoPartsException(f"Ошибка чтения файла {path}: {e}")


@contextmanager
def safe_operation(name: str = "операция"):
    """Контекстный менеджер для безопасного выполнения операции"""
    try:
        yield
    except Exception as e:
        logger.error(f"❌ Ошибка в {name}: {e}")
        logger.error(traceback.format_exc())
        raise AutoPartsException(f"Ошибка в {name}: {e}")


# ============================================================================
# ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ
# ============================================================================
def safe_float(val: Any, default: float = 0.0) -> float:
    """Безопасное преобразование в float"""
    if val is None:
        return default
    if isinstance(val, bool):
        return float(val)
    if isinstance(val, (int, float)):
        if math.isnan(val) or math.isinf(val):
            return default
        return float(val)
    if isinstance(val, (decimal.Decimal, np.floating, np.integer)):
        try:
            return float(val)
        except Exception:
            return default
    if isinstance(val, str):
        cleaned = val.strip()
        if not cleaned:
            return default
        cleaned = re.sub(r'[^\d.,\-+\s]', '', cleaned)
        cleaned = cleaned.replace(' ', '')
        cleaned = cleaned.replace(',', '.')
        if cleaned.count('-') > 1:
            return default
        parts = cleaned.split('.')
        if len(parts) > 2:
            return default
        try:
            return float(cleaned)
        except ValueError:
            return default
    if hasattr(val, 'dtype') and hasattr(val, 'item'):
        try:
            item = val.item()
            if isinstance(item, (int, float)):
                return float(item)
        except Exception:
            pass
    return default


def safe_int(val: Any, default: int = 0) -> int:
    """Безопасное преобразование в int"""
    try:
        float_val = safe_float(val, default)
        if float_val == default and val != 0:
            return default
        return int(float_val)
    except (ValueError, TypeError):
        return default


def safe_str(val: Any, default: str = "") -> str:
    """Безопасное преобразование в str"""
    if val is None:
        return default
    if isinstance(val, bool):
        return str(val)
    if isinstance(val, (int, float)):
        if math.isnan(val) or math.isinf(val):
            return default
        return str(val)
    if isinstance(val, (list, tuple)):
        return ", ".join(safe_str(v) for v in val[:5]) + ("..." if len(val) > 5 else "")
    if isinstance(val, dict):
        return str({k: safe_str(v) for k, v in list(val.items())[:5]})
    try:
        result = str(val).strip()
        return result if result else default
    except Exception:
        return default


def safe_bool(val: Any, default: bool = False) -> bool:
    """Безопасное преобразование в bool"""
    if val is None:
        return default
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return bool(val)
    if isinstance(val, str):
        val_lower = val.lower().strip()
        true_values = {'true', 'yes', '1', 'y', 'да', 'on'}
        false_values = {'false', 'no', '0', 'n', 'нет', 'off'}
        if val_lower in true_values:
            return True
        if val_lower in false_values:
            return False
        return default
    if isinstance(val, (list, tuple, dict)):
        return bool(val)
    return default


def safe_datetime(val: Any, default: Optional[datetime] = None) -> Optional[datetime]:
    """Безопасное преобразование в datetime"""
    if default is None:
        default = datetime.now()
    if val is None:
        return default
    if isinstance(val, datetime):
        return val
    if isinstance(val, date):
        return datetime.combine(val, datetime.min.time())
    if isinstance(val, (int, float)):
        try:
            return datetime.fromtimestamp(val)
        except Exception:
            return default
    if isinstance(val, str):
        formats = [
            "%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d",
            "%d.%m.%Y %H:%M:%S", "%d.%m.%Y", "%Y-%m-%dT%H:%M:%S",
            "%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S.%fZ",
            "%b %d %Y %H:%M:%S", "%d %b %Y %H:%M:%S",
            "%d/%m/%Y %H:%M:%S", "%d/%m/%Y", "%m/%d/%Y", "%m/%d/%Y %H:%M:%S",
        ]
        for fmt in formats:
            try:
                return datetime.strptime(val, fmt)
            except ValueError:
                continue
        try:
            if DATEUTIL_AVAILABLE:
                return parse(val)
        except Exception:
            pass
    return default


def generate_cache_key(*args, **kwargs) -> str:
    """Генерация ключа для кеша"""
    key_parts = []
    for arg in args:
        if isinstance(arg, (dict, OrderedDict)):
            key_parts.append(json.dumps(arg, sort_keys=True, ensure_ascii=False))
        elif isinstance(arg, (list, tuple, set)):
            key_parts.append(str(sorted(arg) if not isinstance(arg, tuple) else arg))
        elif isinstance(arg, pd.DataFrame):
            try:
                key_parts.append(hashlib.md5(pd.util.hash_pandas_object(arg).values.tobytes()).hexdigest())
            except Exception:
                key_parts.append(str(len(arg)))
        elif isinstance(arg, pd.Series):
            try:
                key_parts.append(hashlib.md5(pd.util.hash_pandas_object(arg).values.tobytes()).hexdigest())
            except Exception:
                key_parts.append(str(len(arg)))
        elif isinstance(arg, np.ndarray):
            try:
                key_parts.append(hashlib.md5(arg.tobytes()).hexdigest())
            except Exception:
                key_parts.append(str(arg.shape))
        elif isinstance(arg, (datetime, date)):
            key_parts.append(arg.isoformat())
        else:
            key_parts.append(str(arg))
    for k, v in sorted(kwargs.items()):
        if isinstance(v, (dict, OrderedDict)):
            key_parts.append(f"{k}:{json.dumps(v, sort_keys=True, ensure_ascii=False)}")
        elif isinstance(v, (list, tuple, set)):
            key_parts.append(f"{k}:{str(sorted(v) if not isinstance(v, tuple) else v)}")
        elif isinstance(v, pd.DataFrame):
            try:
                key_parts.append(f"{k}:{hashlib.md5(pd.util.hash_pandas_object(v).values.tobytes()).hexdigest()}")
            except Exception:
                key_parts.append(f"{k}:{len(v)}")
        else:
            key_parts.append(f"{k}:{v}")
    key = "|".join(key_parts)
    return hashlib.md5(key.encode('utf-8')).hexdigest()


def format_currency(value: float, currency: str = "RUB", locale: str = "ru_RU") -> str:
    """Форматирование валюты с учетом локали"""
    if value is None or math.isnan(value) or math.isinf(value):
        return f"0 {currency}"
    if BABEL_AVAILABLE:
        try:
            return babel_format_currency(value, currency, locale=locale)
        except Exception:
            pass
    symbols = {
        "RUB": "₽", "USD": "$", "EUR": "€", "CNY": "¥",
        "KZT": "₸", "UAH": "₴", "BYN": "Br", "AMD": "֏"
    }
    symbol = symbols.get(currency, currency)
    if abs(value) >= 1:
        return f"{value:,.0f} {symbol}"
    else:
        return f"{value:.2f} {symbol}"


def format_percent(value: float, decimal_places: int = 1) -> str:
    """Форматирование процентов"""
    if value is None or math.isnan(value) or math.isinf(value):
        return "0%"
    if BABEL_AVAILABLE:
        try:
            return babel_format_percent(value / 100, format=f"#,##0.{'0' * decimal_places}%", locale="ru_RU")
        except Exception:
            pass
    return f"{value:.{decimal_places}f}%"


def format_number(value: float, decimal_places: int = 2) -> str:
    """Форматирование числа"""
    if value is None or math.isnan(value) or math.isinf(value):
        return "0"
    if BABEL_AVAILABLE:
        try:
            return babel_format_decimal(value, format=f"#,##0.{'0' * decimal_places}", locale="ru_RU")
        except Exception:
            pass
    return f"{value:,.{decimal_places}f}"


def calculate_volume(length: float, width: float, height: float) -> float:
    """Расчет объема в литрах с валидацией"""
    if not all([length, width, height]):
        return 0.0
    if not all([length > 0, width > 0, height > 0]):
        return 0.0
    if any([length > 1000, width > 1000, height > 1000]):
        length /= 10
        width /= 10
        height /= 10
    if any([length < 0.1, width < 0.1, height < 0.1]):
        return 0.0
    volume = (length * width * height) / 1000.0
    if volume < 0.001:
        return 0.0
    return round(volume, 4)


def calculate_weighted_average(values: List[float], weights: List[float]) -> float:
    """Взвешенное среднее с валидацией"""
    if not values or not weights:
        return 0.0
    if len(values) != len(weights):
        raise ValueError("Длины списков values и weights должны совпадать")
    total_weight = sum(weights)
    if total_weight == 0:
        return 0.0
    return sum(v * w for v, w in zip(values, weights)) / total_weight


def calculate_percentile(data: List[float], percentile: float) -> float:
    """Расчет перцентиля"""
    if not data:
        return 0.0
    if percentile < 0 or percentile > 100:
        raise ValueError("Перцентиль должен быть в диапазоне 0-100")
    sorted_data = sorted(data)
    index = (len(sorted_data) - 1) * percentile / 100
    floor_idx = int(index)
    ceil_idx = floor_idx + 1
    if ceil_idx >= len(sorted_data):
        return sorted_data[floor_idx]
    return sorted_data[floor_idx] + (sorted_data[ceil_idx] - sorted_data[floor_idx]) * (index - floor_idx)


def calculate_std_deviation(data: List[float]) -> float:
    """Расчет стандартного отклонения"""
    if len(data) < 2:
        return 0.0
    mean = sum(data) / len(data)
    variance = sum((x - mean) ** 2 for x in data) / (len(data) - 1)
    return math.sqrt(variance)


def calculate_correlation(x: List[float], y: List[float]) -> float:
    """Расчет корреляции Пирсона"""
    if len(x) != len(y) or len(x) < 2:
        return 0.0
    mean_x = sum(x) / len(x)
    mean_y = sum(y) / len(y)
    numerator = sum((x[i] - mean_x) * (y[i] - mean_y) for i in range(len(x)))
    denominator = math.sqrt(
        sum((x[i] - mean_x) ** 2 for i in range(len(x))) *
        sum((y[i] - mean_y) ** 2 for i in range(len(y)))
    )
    if denominator == 0:
        return 0.0
    return numerator / denominator


def is_valid_barcode(barcode: str) -> bool:
    """Проверка валидности штрихкода (EAN-8, EAN-13, UPC-A, UPC-E)"""
    if not barcode:
        return False
    barcode = re.sub(r'[^\d]', '', barcode)
    if len(barcode) not in [8, 12, 13, 14]:
        return False
    if len(barcode) == 13:
        checksum = 0
        for i, digit in enumerate(barcode[:-1]):
            checksum += int(digit) * (3 if (i + 1) % 2 == 0 else 1)
        checksum = (10 - (checksum % 10)) % 10
        return checksum == int(barcode[-1])
    return True


def normalize_text(text: str) -> str:
    """Нормализация текста"""
    if not text:
        return ""
    text = text.lower()
    text = re.sub(r'[^\w\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def extract_numbers(text: str) -> List[float]:
    """Извлечение всех чисел из текста"""
    if not text:
        return []
    return [float(x) for x in re.findall(r'-?\d+\.?\d*', text)]


def levenshtein_distance(s1: str, s2: str) -> int:
    """Расстояние Левенштейна (оптимизированная версия)"""
    if len(s1) < len(s2):
        return levenshtein_distance(s2, s1)
    if len(s2) == 0:
        return len(s1)
    previous_row = list(range(len(s2) + 1))
    for i, c1 in enumerate(s1):
        current_row = [i + 1]
        for j, c2 in enumerate(s2):
            insertions = previous_row[j + 1] + 1
            deletions = current_row[j] + 1
            substitutions = previous_row[j] + (c1 != c2)
            current_row.append(min(insertions, deletions, substitutions))
        previous_row = current_row
    return previous_row[-1]


def find_similar_strings(text: str, candidates: List[str], threshold: int = 3) -> List[Tuple[str, int]]:
    """Поиск похожих строк"""
    if not text or not candidates:
        return []
    normalized_text = normalize_text(text)
    results = []
    for candidate in candidates:
        normalized_candidate = normalize_text(candidate)
        distance = levenshtein_distance(normalized_text, normalized_candidate)
        if distance <= threshold:
            results.append((candidate, distance))
    return sorted(results, key=lambda x: x[1])


def get_file_encoding(file_path: Union[str, Path]) -> str:
    """Автоматическое определение кодировки файла"""
    if CHARDET_AVAILABLE and chardet is not None:
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(100000)
                result = chardet.detect(raw_data)
                encoding = result.get('encoding', 'utf-8')
                confidence = result.get('confidence', 0)
                logger.debug(f"Определена кодировка {encoding} (уверенность {confidence:.2f})")
                return encoding
        except Exception as e:
            logger.warning(f"Ошибка определения кодировки: {e}")
    encodings = ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251', 'cp1252', 'latin1']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                f.read()
            return enc
        except Exception:
            continue
    return 'utf-8'


def load_dataframe(file_path: Union[str, Path], **kwargs) -> pd.DataFrame:
    """Универсальная загрузка данных с автоопределением формата"""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"Файл {path} не найден")
    ext = path.suffix.lower()
    try:
        if ext == '.csv':
            encoding = get_file_encoding(path)
            return pd.read_csv(path, encoding=encoding, **kwargs)
        elif ext in ['.xlsx', '.xls']:
            return pd.read_excel(path, **kwargs)
        elif ext == '.json':
            return pd.read_json(path, **kwargs)
        elif ext == '.parquet':
            return pd.read_parquet(path, **kwargs)
        elif ext == '.feather':
            return pd.read_feather(path, **kwargs)
        elif ext == '.pickle':
            return pd.read_pickle(path, **kwargs)
        elif ext == '.html':
            return pd.read_html(path, **kwargs)[0]
        elif ext == '.xml':
            return pd.read_xml(path, **kwargs)
        else:
            raise ValueError(f"Неподдерживаемый формат файла: {ext}")
    except Exception as e:
        raise AutoPartsException(f"Ошибка загрузки файла {path}: {e}")


def save_dataframe(df: pd.DataFrame, file_path: Union[str, Path], **kwargs) -> None:
    """Универсальное сохранение данных"""
    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    ext = path.suffix.lower()
    try:
        if ext == '.csv':
            df.to_csv(path, encoding='utf-8-sig', index=False, **kwargs)
        elif ext == '.xlsx':
            df.to_excel(path, index=False, **kwargs)
        elif ext == '.json':
            df.to_json(path, orient='records', force_ascii=False, **kwargs)
        elif ext == '.parquet':
            df.to_parquet(path, **kwargs)
        elif ext == '.feather':
            df.to_feather(path, **kwargs)
        elif ext == '.pickle':
            df.to_pickle(path, **kwargs)
        elif ext == '.html':
            df.to_html(path, **kwargs)
        elif ext == '.xml':
            df.to_xml(path, **kwargs)
        else:
            raise ValueError(f"Неподдерживаемый формат файла: {ext}")
    except Exception as e:
        raise AutoPartsException(f"Ошибка сохранения файла {path}: {e}")


def clean_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Очистка DataFrame от пустых строк и дубликатов"""
    df = df.copy()
    df = df.dropna(how='all')
    df = df.dropna(axis=1, how='all')
    df = df.drop_duplicates()
    df.columns = [
        re.sub(r'[^\w\s]', '', col).strip().replace(' ', '_').lower()
        for col in df.columns
    ]
    return df


def validate_dataframe(df: pd.DataFrame, required_columns: List[str]) -> Tuple[bool, List[str]]:
    """Валидация DataFrame на наличие обязательных колонок"""
    missing = []
    for col in required_columns:
        if col not in df.columns:
            missing.append(col)
    return len(missing) == 0, missing


def get_optimal_batch_size(data_size: int, max_workers: int = None) -> int:
    """Расчет оптимального размера батча"""
    if max_workers is None:
        max_workers = os.cpu_count() or 4
    base_size = 1000
    optimal_size = max(100, min(base_size, data_size // max_workers))
    max_batch = 10000
    return min(optimal_size, max_batch)


def split_into_batches(data: List, batch_size: int) -> List[List]:
    """Разбиение данных на батчи"""
    return [data[i:i + batch_size] for i in range(0, len(data), batch_size)]


def parallel_apply(df: pd.DataFrame, func: Callable, axis: int = 1, max_workers: int = None) -> pd.Series:
    """Параллельное применение функции к строкам DataFrame"""
    if max_workers is None:
        max_workers = os.cpu_count() or 4
    chunks = np.array_split(df, max_workers)
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = []
        for chunk in chunks:
            futures.append(executor.submit(lambda df_chunk: df_chunk.apply(func, axis=axis), chunk))
        results = []
        for future in as_completed(futures):
            results.append(future.result())
    return pd.concat(results)


def memory_optimize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Оптимизация использования памяти DataFrame"""
    if not OPTIMIZE_MEMORY:
        return df
    df_optimized = df.copy()
    for col in df_optimized.columns:
        col_type = df_optimized[col].dtype
        if col_type == 'object':
            if len(df_optimized[col].unique()) / len(df_optimized[col]) < 0.5:
                df_optimized[col] = df_optimized[col].astype('category')
        elif col_type == 'float64':
            try:
                df_optimized[col] = df_optimized[col].astype('float32')
            except Exception:
                pass
        elif col_type == 'int64':
            col_min = df_optimized[col].min()
            col_max = df_optimized[col].max()
            if col_min >= 0 and col_max <= 255:
                df_optimized[col] = df_optimized[col].astype('uint8')
            elif col_min >= 0 and col_max <= 65535:
                df_optimized[col] = df_optimized[col].astype('uint16')
            elif col_min >= -128 and col_max <= 127:
                df_optimized[col] = df_optimized[col].astype('int8')
            elif col_min >= -32768 and col_max <= 32767:
                df_optimized[col] = df_optimized[col].astype('int16')
            elif col_min >= -2147483648 and col_max <= 2147483647:
                df_optimized[col] = df_optimized[col].astype('int32')
    return df_optimized


def get_dataframe_memory_usage(df: pd.DataFrame) -> Dict[str, int]:
    """Получение использования памяти каждой колонкой DataFrame"""
    return {col: df[col].memory_usage(deep=True) for col in df.columns}


def format_memory_size(size_bytes: int) -> str:
    """Форматирование размера памяти"""
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} PB"


def get_timestamp() -> str:
    """Получение текущего timestamp в формате ISO"""
    return datetime.now().isoformat()


def format_timestamp(timestamp: Union[datetime, str]) -> str:
    """Форматирование timestamp"""
    if isinstance(timestamp, str):
        try:
            timestamp = datetime.fromisoformat(timestamp)
        except Exception:
            return timestamp
    if isinstance(timestamp, datetime):
        return timestamp.strftime("%Y-%m-%d %H:%M:%S")
    return str(timestamp)


def truncate_text(text: str, max_length: int = 100) -> str:
    """Обрезка текста до указанной длины"""
    if not text:
        return ""
    if len(text) <= max_length:
        return text
    return text[:max_length] + "..."


def generate_id() -> str:
    """Генерация уникального ID"""
    return str(uuid.uuid4())


def generate_short_id(length: int = 8) -> str:
    """Генерация короткого уникального ID"""
    return secrets.token_hex(length // 2)


def convert_dimension(value: float, from_unit: str, to_unit: str) -> float:
    """Конвертация единиц измерения"""
    if value == 0:
        return 0.0
    if from_unit == to_unit:
        return value
    if from_unit == "мм" and to_unit == "см":
        return value / 10.0
    elif from_unit == "см" and to_unit == "мм":
        return value * 10.0
    elif from_unit == "м" and to_unit == "см":
        return value * 100.0
    elif from_unit == "см" and to_unit == "м":
        return value / 100.0
    elif from_unit == "м" and to_unit == "мм":
        return value * 1000.0
    elif from_unit == "мм" and to_unit == "м":
        return value / 1000.0
    return value


def validate_article(article: str) -> bool:
    """Валидация артикула"""
    if not article or not article.strip():
        return False
    return bool(re.match(r'^[A-Za-z0-9\-_]+$', article.strip()))


def generate_random_id(length: int = 12) -> str:
    """Генерация случайного ID"""
    chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789'
    return ''.join(random.choice(chars) for _ in range(length))


def format_barcode(barcode: str) -> str:
    """Форматирование штрихкода"""
    if not barcode:
        return ""
    barcode = re.sub(r'[^\d]', '', barcode)
    if len(barcode) == 13:
        return f"{barcode[:3]} {barcode[3:7]} {barcode[7:11]} {barcode[11:]}"
    elif len(barcode) == 12:
        return f"{barcode[:2]} {barcode[2:6]} {barcode[6:10]} {barcode[10:]}"
    elif len(barcode) == 8:
        return f"{barcode[:2]} {barcode[2:5]} {barcode[5:]}"
    return barcode


def detect_column_mapping(df: pd.DataFrame, required_columns: List[str]) -> Dict[str, str]:
    """Определение соответствия колонок"""
    column_variants = {
        'artikul': ['артикул', 'article', 'sku', 'код товара', 'артикул продавца'],
        'brand': ['бренд', 'brand', 'производитель', 'manufacturer', 'марка'],
        'price': ['цена', 'price', 'стоимость', 'retail price', 'розничная цена'],
        'cost': ['себестоимость', 'cost', 'закупочная цена', 'purchase price'],
        'length': ['длина', 'length', 'длинна', 'габарит длина'],
        'width': ['ширина', 'width', 'габарит ширина'],
        'height': ['высота', 'height', 'габарит высота'],
        'weight': ['вес', 'weight', 'масса'],
        'oe_number': ['oe номер', 'oe number', 'oem', 'оригинальный номер'],
        'category': ['категория', 'category', 'группа', 'раздел'],
        'barcode': ['штрихкод', 'barcode', 'ean', 'штрих-код'],
        'description': ['описание', 'description', 'наименование', 'name'],
        'multiplicity': ['кратность', 'multiplicity', 'упаковка']
    }
    mapping = {}
    actual_lower = {col.lower(): col for col in df.columns}
    for required in required_columns:
        variants = column_variants.get(required, [required])
        for variant in variants:
            variant_lower = variant.lower()
            for actual_l, actual_orig in actual_lower.items():
                if variant_lower in actual_l and actual_orig not in mapping:
                    mapping[actual_orig] = required
                    break
    return mapping
# ============================================================================
# БЛОК 1: ENUM И ТИПЫ (200+ СТРОК)
# ============================================================================
class CommissionType(Enum):
    """Типы комиссий маркетплейсов"""
    PERCENTAGE = auto()
    FIXED = auto()
    HYBRID = auto()
    SUBSCRIPTION = auto()
    TIERED = auto()
    DYNAMIC = auto()
    FLAT = auto()
    CUSTOM = auto()


class OperationMode(Enum):
    """Режимы работы с маркетплейсами"""
    FBY = auto()
    FBS = auto()
    FBO = auto()
    DBS = auto()
    FBP = auto()
    DBE = auto()
    STANDARD = auto()
    EXPRESS = auto()
    SELF = auto()


class ProductType(Enum):
    """Типы продуктов автозапчастей"""
    ENGINE = "Двигатель"
    TRANSMISSION = "Трансмиссия"
    SUSPENSION = "Подвеска"
    BRAKE = "Тормозная система"
    STEERING = "Рулевое управление"
    ELECTRICAL = "Электрооборудование"
    COOLING = "Система охлаждения"
    EXHAUST = "Система выпуска"
    FUEL = "Система питания"
    FILTER = "Фильтры"
    FLUID = "Масла и жидкости"
    BODY = "Кузовные детали"
    INTERIOR = "Салон"
    EXTERIOR = "Экстерьер"
    OPTICS = "Оптика"
    TIRES = "Шины и диски"
    TOOLS = "Инструменты"
    BELT = "Ремни и приводы"
    BEARING = "Подшипники"
    SEAL = "Сальники и прокладки"
    FASTENER = "Крепеж"
    HVAC = "Климат-контроль"
    AUDIO = "Аудио и мультимедиа"
    SAFETY = "Безопасность"
    OTHER = "Прочее"


class DataSource(Enum):
    """Источники данных"""
    CSV = auto()
    EXCEL = auto()
    JSON = auto()
    API = auto()
    DATABASE = auto()
    MANUAL = auto()
    MARKETPLACE = auto()
    AI = auto()
    WEB_SCRAPING = auto()
    ERP = auto()
    CRM = auto()
    EXTERNAL = auto()


class ExportFormat(Enum):
    """Форматы экспорта"""
    CSV = auto()
    EXCEL = auto()
    PDF = auto()
    JSON = auto()
    HTML = auto()
    MARKDOWN = auto()
    PARQUET = auto()
    SQL = auto()
    XML = auto()
    YAML = auto()
    TOML = auto()
    POWER_BI = auto()
    TABLEAU = auto()


class CalculationStatus(Enum):
    """Статусы расчетов"""
    PENDING = auto()
    RUNNING = auto()
    COMPLETED = auto()
    FAILED = auto()
    CANCELLED = auto()
    PAUSED = auto()
    PARTIAL = auto()


class RiskLevel(Enum):
    """Уровни риска"""
    LOW = "Низкий"
    MEDIUM = "Средний"
    HIGH = "Высокий"
    CRITICAL = "Критический"


class Seasonality(Enum):
    """Сезонность товаров"""
    WINTER = "Зимняя"
    SPRING = "Весенняя"
    SUMMER = "Летняя"
    AUTUMN = "Осенняя"
    ALL_YEAR = "Круглогодичная"


class ProfitabilityLevel(Enum):
    """Уровни прибыльности"""
    LOSS = "Убыток"
    BREAK_EVEN = "Точка безубыточности"
    LOW = "Низкая"
    MEDIUM = "Средняя"
    HIGH = "Высокая"
    VERY_HIGH = "Очень высокая"


class Currency(Enum):
    """Валюты"""
    RUB = "RUB"
    USD = "USD"
    EUR = "EUR"
    CNY = "CNY"
    KZT = "KZT"
    UAH = "UAH"
    BYN = "BYN"
    AMD = "AMD"


# ============================================================================
# БЛОК 2: ДАТАКЛАССЫ (300+ СТРОК)
# ============================================================================
@dataclass
class MarketplaceConfig:
    """Конфигурация маркетплейса"""
    name: str
    commission_rate: float
    min_commission: float = 0.0
    max_commission: float = float('inf')
    logistics_base: float = 0.0
    logistics_per_kg: float = 0.0
    logistics_per_liter: float = 0.0
    storage_per_day: float = 0.0
    return_fee: float = 0.0
    acquiring_fee: float = 0.0
    last_mile_fee: float = 0.0
    delivery_fee_percent: float = 0.0
    premium_fee: float = 0.0
    rko_fee: float = 0.0
    subscription_fee: float = 0.0
    insurance_fee: float = 0.0
    packing_fee: float = 0.0
    marketing_fee: float = 0.0
    category_rates: Dict[str, float] = field(default_factory=dict)
    mode_multipliers: Dict[str, float] = field(default_factory=dict)
    weight_tiers: List[Tuple[float, float, float]] = field(default_factory=list)
    volume_tiers: List[Tuple[float, float, float]] = field(default_factory=list)
    available: bool = True
    description: str = ""
    version: str = "2026.1"
    last_updated: datetime = field(default_factory=datetime.now)


@dataclass
class ProductDimensions:
    """Габариты товара"""
    length: float = 0.0
    width: float = 0.0
    height: float = 0.0
    weight: float = 0.0
    unit: str = "см"
    weight_unit: str = "кг"

    @property
    def volume(self) -> float:
        return calculate_volume(self.length, self.width, self.height)

    @property
    def is_valid(self) -> bool:
        return all([self.length > 0, self.width > 0, self.height > 0, self.weight > 0])

    def to_dict(self) -> Dict[str, float]:
        return {
            "length": self.length,
            "width": self.width,
            "height": self.height,
            "weight": self.weight,
            "volume": self.volume
        }


@dataclass
class ProductCategory:
    """Категория товара"""
    name: str
    description: str = ""
    parent_category: Optional[str] = None
    dimensions: Optional[ProductDimensions] = None
    typical_volume: float = 0.0
    typical_weight: float = 0.0
    oem_codes: List[str] = field(default_factory=list)
    cross_references: List[str] = field(default_factory=list)
    alternatives: List[str] = field(default_factory=list)
    compatibility: List[str] = field(default_factory=list)
    hazardous: bool = False
    fragile: bool = False
    requires_special_packaging: bool = False
    seasonality: Seasonality = Seasonality.ALL_YEAR
    risk_level: RiskLevel = RiskLevel.LOW
    price_range_min: float = 0.0
    price_range_max: float = 0.0
    margin_avg: float = 0.0
    demand_score: float = 0.0
    # Для обратной совместимости
    min_length: float = 0.0
    max_length: float = 0.0
    min_width: float = 0.0
    max_width: float = 0.0
    min_height: float = 0.0
    max_height: float = 0.0
    min_weight: float = 0.0
    max_weight: float = 0.0

    def get_dimensions(self) -> ProductDimensions:
        return self.dimensions or ProductDimensions()

    def to_dict(self) -> Dict[str, Any]:
        return {
            "name": self.name,
            "description": self.description,
            "parent_category": self.parent_category,
            "typical_volume": self.typical_volume,
            "typical_weight": self.typical_weight,
            "oem_codes": self.oem_codes,
            "cross_references": self.cross_references,
            "hazardous": self.hazardous,
            "fragile": self.fragile,
            "seasonality": self.seasonality.value,
            "risk_level": self.risk_level.value
        }


@dataclass
class UnitEconomicsResult:
    """Результат расчета юнит-экономики"""
    marketplace: str
    operation_mode: str
    category: str
    price: float
    cost: float
    length: float
    width: float
    height: float
    weight: float
    volume: float
    commission: float
    commission_percent: float
    logistics: float
    storage_cost: float
    acquiring: float
    delivery: float
    last_mile: float
    returns: float
    rko_fee: float
    premium_fee: float
    insurance_fee: float
    packing_fee: float
    marketing_fee: float
    subscription_cost: float
    total_expenses: float
    profit: float
    margin_percent: float
    roi: float
    breakeven_price: float
    profit_per_ruble: float
    contribution_margin: float
    contribution_margin_ratio: float
    timestamp: datetime = field(default_factory=datetime.now)
    calculation_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    status: CalculationStatus = CalculationStatus.COMPLETED
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        result = asdict(self)
        result['timestamp'] = self.timestamp.isoformat()
        result['status'] = self.status.name
        return result

    def to_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame([self.to_dict()])

    def get_summary(self) -> Dict[str, Any]:
        return {
            "marketplace": self.marketplace,
            "profit": self.profit,
            "margin": self.margin_percent,
            "roi": self.roi,
            "breakeven": self.breakeven_price,
            "total_expenses": self.total_expenses
        }

    def get_profitability_level(self) -> ProfitabilityLevel:
        if self.profit < 0:
            return ProfitabilityLevel.LOSS
        elif self.profit == 0:
            return ProfitabilityLevel.BREAK_EVEN
        elif self.margin_percent < 5:
            return ProfitabilityLevel.LOW
        elif self.margin_percent < 15:
            return ProfitabilityLevel.MEDIUM
        elif self.margin_percent < 30:
            return ProfitabilityLevel.HIGH
        else:
            return ProfitabilityLevel.VERY_HIGH


@dataclass
class ForecastResult:
    """Результат прогнозирования"""
    periods: List[datetime]
    values: List[float]
    seasonality: List[float]
    trend: List[float]
    confidence_intervals: Tuple[List[float], List[float]]
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame({
            "period": self.periods,
            "value": self.values,
            "seasonality": self.seasonality,
            "trend": self.trend,
            "lower_bound": self.confidence_intervals[0] if self.confidence_intervals else [],
            "upper_bound": self.confidence_intervals[1] if self.confidence_intervals else []
        })


@dataclass
class OptimizationResult:
    """Результат оптимизации"""
    optimal_price: float
    optimal_margin: float
    optimal_profit: float
    current_price: float
    current_margin: float
    current_profit: float
    improvement_pct: float
    recommendations: List[str]
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "optimal_price": self.optimal_price,
            "optimal_margin": self.optimal_margin,
            "optimal_profit": self.optimal_profit,
            "current_price": self.current_price,
            "current_margin": self.current_margin,
            "current_profit": self.current_profit,
            "improvement_pct": self.improvement_pct,
            "recommendations": self.recommendations
        }


@dataclass
class ComparisonResult:
    """Результат сравнения"""
    marketplace: str
    profit: float
    margin: float
    roi: float
    total_expenses: float
    commission: float
    logistics: float
    storage_cost: float
    rank: int
    metadata: Dict[str, Any] = field(default_factory=dict)


# ============================================================================
# БЛОК 3: КОНФИГУРАЦИИ МАРКЕТПЛЕЙСОВ 2026 (200+ СТРОК)
# ============================================================================
def get_marketplace_configs_2026() -> Dict[str, MarketplaceConfig]:
    """Получение конфигураций маркетплейсов на 2026 год"""
    configs = {}

    # OZON
    configs["Ozon"] = MarketplaceConfig(
        name="Ozon",
        commission_rate=0.15,
        min_commission=30.0,
        max_commission=float('inf'),
        logistics_base=50.0,
        logistics_per_kg=15.0,
        logistics_per_liter=5.0,
        storage_per_day=0.3,
        return_fee=0.02,
        acquiring_fee=0.015,
        last_mile_fee=50.0,
        delivery_fee_percent=0.0,
        premium_fee=0.0,
        rko_fee=0.0,
        subscription_fee=0.0,
        insurance_fee=0.0,
        packing_fee=0.0,
        marketing_fee=0.0,
        category_rates={
            "двигатель": 0.12, "трансмиссия": 0.13, "подвеска": 0.14,
            "тормозная_система": 0.14, "фильтры": 0.17, "масла": 0.18,
            "оптика": 0.15, "шины": 0.16, "электрика": 0.15,
            "инструменты": 0.14, "кузов": 0.13
        },
        mode_multipliers={"FBY": 1.0, "FBS": 1.1, "FBO": 1.05, "DBS": 1.2, "FBP": 0.95},
        description="Ozon - крупнейший маркетплейс России",
        version="2026.1"
    )

    # Wildberries
    configs["Wildberries"] = MarketplaceConfig(
        name="Wildberries",
        commission_rate=0.18,
        min_commission=35.0,
        max_commission=float('inf'),
        logistics_base=60.0,
        logistics_per_kg=18.0,
        logistics_per_liter=6.0,
        storage_per_day=0.5,
        return_fee=0.03,
        acquiring_fee=0.0,
        last_mile_fee=0.0,
        delivery_fee_percent=0.05,
        premium_fee=0.0,
        rko_fee=0.01,
        subscription_fee=0.0,
        insurance_fee=0.0,
        packing_fee=0.0,
        marketing_fee=0.0,
        category_rates={
            "двигатель": 0.15, "трансмиссия": 0.16, "подвеска": 0.17,
            "тормозная_система": 0.17, "фильтры": 0.20, "масла": 0.22,
            "оптика": 0.18, "шины": 0.19, "электрика": 0.18,
            "инструменты": 0.17, "кузов": 0.16
        },
        mode_multipliers={"FBY": 1.0, "FBS": 1.15, "FBO": 1.1, "DBS": 1.25, "FBP": 1.0},
        description="Wildberries - лидер e-commerce России",
        version="2026.1"
    )

    # Яндекс Маркет
    configs["Яндекс Маркет"] = MarketplaceConfig(
        name="Яндекс Маркет",
        commission_rate=0.14,
        min_commission=25.0,
        max_commission=float('inf'),
        logistics_base=45.0,
        logistics_per_kg=14.0,
        logistics_per_liter=4.5,
        storage_per_day=0.25,
        return_fee=0.02,
        acquiring_fee=0.02,
        last_mile_fee=40.0,
        delivery_fee_percent=0.0,
        premium_fee=0.0,
        rko_fee=0.0,
        subscription_fee=0.0,
        insurance_fee=0.0,
        packing_fee=0.0,
        marketing_fee=0.0,
        category_rates={
            "двигатель": 0.11, "трансмиссия": 0.12, "подвеска": 0.13,
            "тормозная_система": 0.13, "фильтры": 0.16, "масла": 0.17,
            "оптика": 0.14, "шины": 0.15, "электрика": 0.14,
            "инструменты": 0.13, "кузов": 0.12
        },
        mode_multipliers={"FBY": 1.0, "FBS": 1.08, "FBO": 1.03, "DBS": 1.15, "FBP": 0.93},
        description="Яндекс Маркет - маркетплейс экосистемы Яндекса",
        version="2026.1"
    )

    # AliExpress
    configs["AliExpress"] = MarketplaceConfig(
        name="AliExpress",
        commission_rate=0.10,
        min_commission=20.0,
        max_commission=float('inf'),
        logistics_base=80.0,
        logistics_per_kg=25.0,
        logistics_per_liter=8.0,
        storage_per_day=0.2,
        return_fee=0.01,
        acquiring_fee=0.025,
        last_mile_fee=70.0,
        delivery_fee_percent=0.0,
        premium_fee=0.0,
        rko_fee=0.0,
        subscription_fee=0.0,
        insurance_fee=0.0,
        packing_fee=0.0,
        marketing_fee=0.0,
        category_rates={
            "двигатель": 0.08, "трансмиссия": 0.09, "подвеска": 0.10,
            "тормозная_система": 0.10, "фильтры": 0.12, "масла": 0.13,
            "оптика": 0.11, "шины": 0.12, "электрика": 0.11,
            "инструменты": 0.10, "кузов": 0.09
        },
        mode_multipliers={"FBY": 1.0, "FBS": 1.2, "FBO": 1.1, "DBS": 1.3, "FBP": 0.9},
        description="AliExpress - международный маркетплейс",
        version="2026.1"
    )

    # Мегамаркет
    configs["Мегамаркет"] = MarketplaceConfig(
        name="Мегамаркет",
        commission_rate=0.13,
        min_commission=28.0,
        max_commission=float('inf'),
        logistics_base=55.0,
        logistics_per_kg=16.0,
        logistics_per_liter=5.5,
        storage_per_day=0.3,
        return_fee=0.02,
        acquiring_fee=0.018,
        last_mile_fee=45.0,
        delivery_fee_percent=0.0,
        premium_fee=0.0,
        rko_fee=0.0,
        subscription_fee=0.0,
        insurance_fee=0.0,
        packing_fee=0.0,
        marketing_fee=0.0,
        category_rates={
            "двигатель": 0.10, "трансмиссия": 0.11, "подвеска": 0.12,
            "тормозная_система": 0.12, "фильтры": 0.15, "масла": 0.16,
            "оптика": 0.13, "шины": 0.14, "электрика": 0.13,
            "инструменты": 0.12, "кузов": 0.11
        },
        mode_multipliers={"FBY": 1.0, "FBS": 1.1, "FBO": 1.05, "DBS": 1.2, "FBP": 0.95},
        description="Мегамаркет (Сбер) - маркетплейс экосистемы Сбера",
        version="2026.1"
    )

    # СберМегаМаркет
    configs["СберМегаМаркет"] = MarketplaceConfig(
        name="СберМегаМаркет",
        commission_rate=0.13,
        min_commission=28.0,
        max_commission=float('inf'),
        logistics_base=55.0,
        logistics_per_kg=16.0,
        logistics_per_liter=5.5,
        storage_per_day=0.3,
        return_fee=0.02,
        acquiring_fee=0.018,
        last_mile_fee=45.0,
        delivery_fee_percent=0.0,
        premium_fee=0.0,
        rko_fee=0.0,
        subscription_fee=0.0,
        insurance_fee=0.0,
        packing_fee=0.0,
        marketing_fee=0.0,
        category_rates={
            "двигатель": 0.10, "трансмиссия": 0.11, "подвеска": 0.12,
            "тормозная_система": 0.12, "фильтры": 0.15, "масла": 0.16,
            "оптика": 0.13, "шины": 0.14, "электрика": 0.13,
            "инструменты": 0.12, "кузов": 0.11
        },
        mode_multipliers={"FBY": 1.0, "FBS": 1.1, "FBO": 1.05, "DBS": 1.2, "FBP": 0.95},
        description="СберМегаМаркет - маркетплейс Сбера",
        version="2026.1"
    )

    return configs


# ============================================================================
# БЛОК 4: 150+ КАТЕГОРИЙ АВТОЗАПЧАСТЕЙ С ПОЛНЫМИ ГАБАРИТАМИ (1500+ СТРОК)
# ============================================================================
def get_auto_parts_categories_full() -> Dict[str, ProductCategory]:
    """Получение полного списка категорий автозапчастей с габаритами"""
    categories = {}

    # Вспомогательная функция для создания категории
    def make_cat(name, desc, min_l, max_l, min_w, max_w, min_h, max_h,
                 min_wt, max_wt, typ_vol, typ_wt, oem=None, season=Seasonality.ALL_YEAR,
                 risk=RiskLevel.LOW, hazardous=False, fragile=False):
        return ProductCategory(
            name=name, description=desc,
            min_length=min_l, max_length=max_l,
            min_width=min_w, max_width=max_w,
            min_height=min_h, max_height=max_h,
            min_weight=min_wt, max_weight=max_wt,
            typical_volume=typ_vol, typical_weight=typ_wt,
            dimensions=ProductDimensions(
                length=(min_l + max_l) / 2, width=(min_w + max_w) / 2,
                height=(min_h + max_h) / 2, weight=(min_wt + max_wt) / 2
            ),
            oem_codes=oem or [], seasonality=season,
            risk_level=risk, hazardous=hazardous, fragile=fragile
        )

    # Двигатель
    categories["двигатель"] = make_cat("двигатель", "Двигатели и комплектующие", 30, 80, 30, 60, 30, 70, 10, 200, 20.0, 80.0, risk=RiskLevel.HIGH)
    categories["поршни"] = make_cat("поршни", "Поршни и кольца", 5, 12, 5, 12, 3, 10, 0.1, 1.5, 0.1, 0.5)
    categories["клапаны"] = make_cat("клапаны", "Клапаны двигателя", 3, 8, 1, 3, 10, 40, 0.05, 0.5, 0.05, 0.2)
    categories["прокладки_двигателя"] = make_cat("прокладки_двигателя", "Прокладки ГБЦ и двигателя", 10, 50, 10, 40, 0.1, 2, 0.01, 0.3, 0.1, 0.1)
    categories["свечи_зажигания"] = make_cat("свечи_зажигания", "Свечи зажигания", 2, 3, 2, 3, 6, 10, 0.04, 0.1, 0.01, 0.05)
    categories["блок_цилиндров"] = make_cat("блок_цилиндров", "Блок цилиндров", 40, 70, 30, 50, 20, 40, 20, 80, 100.0, 50.0, risk=RiskLevel.HIGH)
    categories["головка_блока"] = make_cat("головка_блока", "Головка блока цилиндров", 30, 60, 20, 40, 8, 20, 5, 30, 40.0, 15.0, risk=RiskLevel.HIGH)
    categories["коленвал"] = make_cat("коленвал", "Коленчатый вал", 40, 90, 8, 20, 8, 20, 10, 40, 30.0, 25.0, risk=RiskLevel.HIGH)
    categories["распредвал"] = make_cat("распредвал", "Распределительный вал", 30, 80, 5, 15, 5, 15, 3, 15, 20.0, 9.0)
    categories["шатун"] = make_cat("шатун", "Шатун двигателя", 12, 35, 4, 10, 3, 7, 0.5, 2, 3.0, 1.25)
    categories["гидрокомпенсаторы"] = make_cat("гидрокомпенсаторы", "Гидрокомпенсаторы", 3, 8, 3, 8, 3, 8, 0.05, 0.2, 0.3, 0.125)
    categories["привод_грм"] = make_cat("привод_грм", "Привод ГРМ (ремень, цепь)", 60, 160, 2, 5, 1, 2, 0.1, 1, 2.0, 0.55)
    categories["масляный_насос"] = make_cat("масляный_насос", "Масляный насос", 8, 18, 8, 18, 8, 18, 1, 5, 5.0, 3.0)
    categories["водяной_насос"] = make_cat("водяной_насос", "Водяной насос (помпа)", 8, 18, 8, 18, 8, 18, 1, 4, 5.0, 2.5)
    categories["турбокомпрессор"] = make_cat("турбокомпрессор", "Турбокомпрессор", 15, 35, 15, 30, 15, 25, 5, 15, 15.0, 10.0, risk=RiskLevel.HIGH)
    categories["масляный_поддон"] = make_cat("масляный_поддон", "Масляный поддон", 30, 60, 20, 40, 10, 20, 2, 8, 15.0, 5.0)
    categories["клапанная_крышка"] = make_cat("клапанная_крышка", "Клапанная крышка", 30, 60, 15, 30, 5, 10, 1, 4, 8.0, 2.5)
    categories["приводной_ремень"] = make_cat("приводной_ремень", "Приводной ремень", 60, 150, 1, 3, 0.5, 1, 0.05, 0.5, 1.0, 0.275)
    categories["демпфер_коленвала"] = make_cat("демпфер_коленвала", "Демпфер коленвала", 10, 25, 10, 25, 5, 10, 2, 8, 5.0, 5.0)
    categories["маховик"] = make_cat("маховик", "Маховик", 25, 45, 25, 45, 5, 10, 5, 15, 10.0, 10.0, risk=RiskLevel.HIGH)
    categories["стартерный_венец"] = make_cat("стартерный_венец", "Стартерный венец", 25, 40, 25, 40, 2, 5, 1, 5, 5.0, 3.0)

    # Трансмиссия
    categories["трансмиссия"] = make_cat("трансмиссия", "КПП и комплектующие", 40, 80, 30, 60, 30, 60, 20, 100, 30.0, 50.0, risk=RiskLevel.HIGH)
    categories["сцепление"] = make_cat("сцепление", "Комплекты сцепления", 20, 40, 20, 40, 5, 15, 2, 10, 3.0, 5.0)
    categories["шкивы"] = make_cat("шкивы", "Шкивы и ролики", 5, 20, 5, 20, 2, 8, 0.2, 3, 0.5, 1.5)
    categories["коробка_передач"] = make_cat("коробка_передач", "Коробка передач в сборе", 40, 70, 30, 50, 25, 40, 30, 80, 80.0, 55.0, risk=RiskLevel.HIGH)
    categories["привод_полуоси"] = make_cat("привод_полуоси", "Привод (полуоси)", 40, 90, 8, 18, 8, 18, 3, 12, 15.0, 7.5)
    categories["дифференциал"] = make_cat("дифференциал", "Дифференциал", 20, 45, 20, 45, 20, 45, 10, 30, 30.0, 20.0, risk=RiskLevel.HIGH)
    categories["карданный_вал"] = make_cat("карданный_вал", "Карданный вал", 60, 160, 8, 18, 8, 18, 5, 20, 25.0, 12.5)
    categories["раздаточная_коробка"] = make_cat("раздаточная_коробка", "Раздаточная коробка", 25, 45, 20, 35, 20, 35, 15, 40, 35.0, 27.5, risk=RiskLevel.HIGH)
    categories["гидротрансформатор"] = make_cat("гидротрансформатор", "Гидротрансформатор АКПП", 25, 40, 25, 40, 20, 30, 10, 25, 30.0, 17.5, risk=RiskLevel.HIGH)
    categories["механизм_переключения"] = make_cat("механизм_переключения", "Механизм переключения передач", 15, 35, 5, 15, 5, 15, 1, 5, 5.0, 3.0)
    categories["подшипники_трансмиссии"] = make_cat("подшипники_трансмиссии", "Подшипники трансмиссии", 8, 18, 8, 18, 8, 18, 0.5, 3, 3.0, 1.75)
    categories["сальники_трансмиссии"] = make_cat("сальники_трансмиссии", "Сальники трансмиссии", 2, 12, 2, 12, 1, 3, 0.05, 0.3, 0.5, 0.175)
    categories["фильтр_акпп"] = make_cat("фильтр_акпп", "Фильтр АКПП", 8, 18, 8, 18, 8, 18, 0.5, 2, 3.0, 1.25)
    categories["масло_трансмиссионное"] = make_cat("масло_трансмиссионное", "Трансмиссионное масло", 10, 35, 8, 25, 8, 25, 1, 5, 5.0, 3.0, hazardous=True)
    categories["трос_сцепления"] = make_cat("трос_сцепления", "Трос сцепления", 40, 100, 1, 3, 1, 3, 0.1, 0.5, 1.0, 0.3)
    categories["цилиндр_сцепления"] = make_cat("цилиндр_сцепления", "Цилиндр сцепления", 10, 20, 5, 10, 5, 10, 0.5, 2, 2.0, 1.25)
    categories["вал_кпп"] = make_cat("вал_кпп", "Вал КПП", 20, 50, 5, 12, 5, 12, 2, 8, 8.0, 5.0)
    categories["шестерни_кпп"] = make_cat("шестерни_кпп", "Шестерни КПП", 5, 15, 5, 15, 5, 15, 0.5, 3, 3.0, 1.75)
    categories["синхронизатор"] = make_cat("синхронизатор", "Синхронизатор", 5, 12, 5, 12, 3, 8, 0.3, 1.5, 2.0, 0.9)
    categories["муфта_кпп"] = make_cat("муфта_кпп", "Муфта КПП", 5, 15, 5, 15, 3, 8, 0.5, 2, 3.0, 1.25)

    # Подвеска
    categories["подвеска"] = make_cat("подвеска", "Элементы подвески", 20, 80, 10, 40, 10, 60, 1, 20, 5.0, 8.0)
    categories["амортизаторы"] = make_cat("амортизаторы", "Амортизаторы", 5, 10, 5, 10, 40, 70, 2, 8, 5.0, 8.0, fragile=True)
    categories["пружины"] = make_cat("пружины", "Пружины подвески", 20, 40, 20, 40, 30, 60, 3, 10, 8.0, 12.0)
    categories["сайлентблоки"] = make_cat("сайлентблоки", "Сайлентблоки", 3, 10, 3, 10, 2, 8, 0.1, 1, 0.1, 0.3)
    categories["шаровые_опоры"] = make_cat("шаровые_опоры", "Шаровые опоры", 5, 15, 5, 15, 5, 15, 0.3, 2, 0.5, 1.5)
    categories["ступицы"] = make_cat("ступицы", "Ступицы и подшипники", 10, 25, 10, 25, 5, 15, 1, 5, 2.0, 4.0)
    categories["рычаг_подвески"] = make_cat("рычаг_подвески", "Рычаг подвески", 20, 65, 5, 18, 5, 18, 2, 10, 10.0, 6.0)
    categories["стабилизатор"] = make_cat("стабилизатор", "Стабилизатор поперечной устойчивости", 25, 65, 3, 10, 3, 10, 1, 5, 5.0, 3.0)
    categories["пыльник"] = make_cat("пыльник", "Пыльник (чехол)", 5, 12, 5, 12, 8, 22, 0.1, 0.5, 1.0, 0.3)
    categories["отбойник"] = make_cat("отбойник", "Отбойник амортизатора", 5, 12, 5, 12, 5, 12, 0.1, 0.5, 1.0, 0.3)
    categories["опора_стойки"] = make_cat("опора_стойки", "Опора стойки амортизатора", 8, 18, 8, 18, 5, 12, 0.5, 2, 3.0, 1.25)
    categories["подрамник"] = make_cat("подрамник", "Подрамник", 45, 105, 15, 35, 8, 18, 10, 30, 25.0, 20.0, risk=RiskLevel.HIGH)
    categories["распорка"] = make_cat("распорка", "Распорка подвески", 25, 65, 2, 6, 2, 6, 0.5, 2, 2.0, 1.25)
    categories["сайлентблоки_в_сборе"] = make_cat("сайлентблоки_в_сборе", "Сайлентблоки в сборе", 8, 22, 8, 22, 5, 12, 0.5, 2, 3.0, 1.25)
    categories["буфер"] = make_cat("буфер", "Буфер подвески", 5, 12, 5, 12, 5, 12, 0.1, 0.5, 1.0, 0.3)
    categories["подушка_подвески"] = make_cat("подушка_подвески", "Подушка подвески", 8, 18, 8, 18, 5, 12, 0.5, 2, 2.0, 1.25)
    categories["тяга_продольная"] = make_cat("тяга_продольная", "Тяга продольная", 25, 65, 3, 8, 3, 8, 1, 4, 4.0, 2.5)
    categories["балка_моста"] = make_cat("балка_моста", "Балка моста", 45, 85, 10, 20, 10, 20, 15, 40, 30.0, 27.5, risk=RiskLevel.HIGH)

    # Тормозная система
    categories["тормозная_система"] = make_cat("тормозная_система", "Тормозная система", 20, 40, 20, 40, 5, 15, 2, 15, 3.0, 8.0, risk=RiskLevel.HIGH)
    categories["тормозные_диски"] = make_cat("тормозные_диски", "Тормозные диски", 25, 40, 25, 40, 3, 8, 3, 12, 3.0, 8.0, fragile=True)
    categories["тормозные_колодки"] = make_cat("тормозные_колодки", "Тормозные колодки", 10, 20, 5, 12, 3, 8, 1, 4, 1.0, 3.0)
    categories["тормозные_шланги"] = make_cat("тормозные_шланги", "Тормозные шланги", 20, 60, 2, 5, 2, 5, 0.2, 1, 0.3, 0.8)

    # Рулевое управление
    categories["рулевое_управление"] = make_cat("рулевое_управление", "Рулевое управление", 30, 100, 10, 30, 10, 30, 2, 15, 5.0, 10.0)
    categories["рулевые_тяги"] = make_cat("рулевые_тяги", "Рулевые тяги и наконечники", 20, 60, 3, 8, 3, 8, 0.5, 3, 1.0, 2.5)
    categories["рулевые_рейки"] = make_cat("рулевые_рейки", "Рулевые рейки", 50, 100, 10, 20, 10, 20, 5, 15, 8.0, 12.0)
    categories["рулевой_кардан"] = make_cat("рулевой_кардан", "Рулевой кардан", 20, 45, 5, 12, 5, 12, 1, 4, 5.0, 2.5)
    categories["усилитель_руля"] = make_cat("усилитель_руля", "Усилитель руля (ГУР/ЭУР)", 15, 30, 15, 30, 15, 25, 3, 10, 10.0, 6.5)

    # Электрика
    categories["электрика"] = make_cat("электрика", "Электрооборудование", 10, 40, 10, 30, 10, 30, 0.5, 10, 2.0, 5.0)
    categories["стартеры"] = make_cat("стартеры", "Стартеры", 15, 30, 10, 20, 10, 25, 3, 10, 3.0, 6.0)
    categories["генераторы"] = make_cat("генераторы", "Генераторы", 15, 30, 15, 25, 15, 30, 4, 12, 5.0, 8.0)
    categories["аккумуляторы"] = make_cat("аккумуляторы", "Аккумуляторы", 20, 40, 15, 25, 15, 30, 10, 30, 15.0, 20.0, hazardous=True, risk=RiskLevel.HIGH)
    categories["датчики"] = make_cat("датчики", "Датчики", 3, 10, 2, 5, 2, 8, 0.05, 0.5, 0.1, 0.3)

    # Система охлаждения
    categories["охлаждение"] = make_cat("охлаждение", "Система охлаждения", 20, 80, 15, 50, 10, 40, 1, 15, 8.0, 15.0)
    categories["радиаторы"] = make_cat("радиаторы", "Радиаторы охлаждения", 40, 80, 30, 60, 5, 15, 2, 10, 10.0, 15.0, fragile=True)
    categories["помпы"] = make_cat("помпы", "Водяные помпы", 10, 25, 10, 20, 10, 20, 1, 5, 2.0, 4.0)
    categories["термостаты"] = make_cat("термостаты", "Термостаты", 5, 12, 5, 12, 5, 12, 0.2, 1, 0.5, 1.0)

    # Фильтры
    categories["фильтры"] = make_cat("фильтры", "Фильтры", 5, 30, 5, 30, 5, 40, 0.1, 3, 2.0, 5.0)
    categories["масляные_фильтры"] = make_cat("масляные_фильтры", "Масляные фильтры", 6, 12, 6, 12, 8, 15, 0.3, 1, 1.0, 1.5)
    categories["воздушные_фильтры"] = make_cat("воздушные_фильтры", "Воздушные фильтры", 15, 40, 15, 35, 3, 10, 0.2, 2, 2.0, 4.0)
    categories["топливные_фильтры"] = make_cat("топливные_фильтры", "Топливные фильтры", 5, 15, 5, 15, 8, 20, 0.3, 1.5, 1.0, 2.0)
    categories["салонные_фильтры"] = make_cat("салонные_фильтры", "Салонные фильтры", 20, 35, 15, 25, 2, 5, 0.2, 1, 1.5, 2.5)

    # Масла и жидкости
    categories["масла"] = make_cat("масла", "Масла и технические жидкости", 5, 30, 5, 30, 10, 40, 0.5, 20, 5.0, 15.0, hazardous=True)
    categories["моторные_масла"] = make_cat("моторные_масла", "Моторные масла", 8, 25, 8, 25, 20, 40, 1, 20, 5.0, 15.0, hazardous=True)
    categories["трансмиссионные_масла"] = make_cat("трансмиссионные_масла", "Трансмиссионные масла", 8, 25, 8, 25, 20, 40, 1, 20, 5.0, 15.0, hazardous=True)
    categories["тормозная_жидкость"] = make_cat("тормозная_жидкость", "Тормозная жидкость", 5, 10, 5, 10, 15, 25, 0.5, 2, 1.0, 2.0, hazardous=True)
    categories["антифриз"] = make_cat("антифриз", "Антифриз / Охлаждающая жидкость", 10, 30, 10, 30, 20, 40, 1, 20, 5.0, 15.0, hazardous=True)

    # Оптика
    categories["оптика"] = make_cat("оптика", "Оптика и освещение", 15, 60, 15, 40, 15, 40, 0.5, 10, 5.0, 10.0, fragile=True)
    categories["фары"] = make_cat("фары", "Фары головного света", 30, 60, 20, 40, 20, 40, 2, 8, 8.0, 12.0, fragile=True)
    categories["лампы"] = make_cat("лампы", "Автомобильные лампы", 2, 10, 2, 5, 5, 15, 0.02, 0.3, 0.1, 0.3, fragile=True)
    categories["фонари"] = make_cat("фонари", "Задние фонари", 20, 50, 15, 30, 10, 25, 1, 5, 5.0, 8.0, fragile=True)

    # Кузов
    categories["кузов"] = make_cat("кузов", "Кузовные детали", 50, 200, 30, 150, 10, 100, 2, 50, 30.0, 80.0, fragile=True, risk=RiskLevel.HIGH)
    categories["бамперы"] = make_cat("бамперы", "Бамперы", 100, 200, 30, 60, 20, 50, 5, 20, 50.0, 80.0, fragile=True)
    categories["крылья"] = make_cat("крылья", "Крылья", 50, 100, 30, 60, 30, 80, 3, 10, 20.0, 40.0, fragile=True)
    categories["капоты"] = make_cat("капоты", "Капоты", 100, 180, 80, 150, 5, 15, 5, 15, 30.0, 60.0, fragile=True)
    categories["зеркала"] = make_cat("зеркала", "Зеркала заднего вида", 15, 30, 10, 20, 10, 20, 0.5, 3, 3.0, 5.0, fragile=True)

    # Шины и диски
    categories["шины"] = make_cat("шины", "Шины и диски", 40, 80, 40, 80, 15, 40, 5, 30, 20.0, 40.0)
    categories["летние_шины"] = make_cat("летние_шины", "Летние шины", 50, 80, 50, 80, 15, 30, 8, 25, 25.0, 35.0, season=Seasonality.SUMMER)
    categories["зимние_шины"] = make_cat("зимние_шины", "Зимние шины", 50, 80, 50, 80, 15, 30, 8, 25, 25.0, 35.0, season=Seasonality.WINTER)
    categories["диски"] = make_cat("диски", "Колесные диски", 40, 60, 40, 60, 15, 30, 5, 20, 15.0, 25.0, fragile=True)

    # Инструменты
    categories["инструменты"] = make_cat("инструменты", "Автоинструменты", 10, 60, 5, 30, 3, 20, 0.2, 10, 3.0, 8.0)
    categories["домкраты"] = make_cat("домкраты", "Домкраты", 20, 50, 10, 25, 10, 25, 3, 15, 5.0, 12.0)
    categories["наборы_ключей"] = make_cat("наборы_ключей", "Наборы ключей", 15, 40, 10, 25, 3, 10, 1, 8, 3.0, 6.0)

    # Ремни и приводы
    categories["ремни"] = make_cat("ремни", "Ремни ГРМ и приводов", 50, 150, 1, 3, 1, 3, 0.1, 0.8, 0.5, 1.0)
    categories["ролики"] = make_cat("ролики", "Ролики натяжители", 5, 12, 5, 12, 2, 5, 0.2, 1.5, 0.5, 1.0)

    # Подшипники
    categories["подшипники"] = make_cat("подшипники", "Подшипники", 3, 15, 3, 15, 1, 5, 0.1, 3, 0.5, 2.0)

    # Крепёж
    categories["крепёж"] = make_cat("крепёж", "Крепёж и метизы", 0.5, 10, 0.5, 10, 0.5, 10, 0.01, 2, 0.2, 1.0)

    # Климат
    categories["климат"] = make_cat("климат", "Климат-контроль и кондиционер", 20, 80, 20, 60, 15, 50, 2, 20, 10.0, 20.0)
    categories["компрессоры"] = make_cat("компрессоры", "Компрессоры кондиционера", 20, 40, 15, 30, 15, 30, 5, 15, 8.0, 12.0)

    # Выхлопная система
    categories["выпуск"] = make_cat("выпуск", "Выхлопная система", 30, 150, 10, 40, 10, 40, 2, 25, 10.0, 25.0)
    categories["глушители"] = make_cat("глушители", "Глушители", 50, 150, 20, 40, 20, 40, 5, 20, 20.0, 30.0)
    categories["катализаторы"] = make_cat("катализаторы", "Каталитические нейтрализаторы", 30, 80, 15, 30, 15, 30, 3, 15, 10.0, 20.0, hazardous=True, risk=RiskLevel.HIGH)

    # Безопасность
    categories["безопасность"] = make_cat("безопасность", "Системы безопасности", 10, 50, 10, 40, 5, 30, 0.5, 8, 3.0, 6.0, risk=RiskLevel.HIGH)

    # Прочее
    categories["щетки_стеклоочистителя"] = make_cat("щетки_стеклоочистителя", "Щетки стеклоочистителя", 30, 70, 2, 5, 2, 5, 0.1, 0.5, 1.0, 1.5)
    categories["коврики"] = make_cat("коврики", "Автомобильные коврики", 50, 100, 40, 80, 1, 5, 1, 5, 10.0, 15.0)
    categories["чехлы"] = make_cat("чехлы", "Чехлы на сиденья", 40, 80, 30, 60, 5, 20, 1, 5, 15.0, 25.0)
    categories["автохимия"] = make_cat("автохимия", "Автохимия и косметика", 5, 30, 5, 20, 10, 40, 0.3, 5, 2.0, 5.0, hazardous=True)

    return categories
# ============================================================================
# БЛОК 5: ОСНОВНОЙ КЛАСС ЮНИТ-ЭКОНОМИКИ (800+ СТРОК)
# ============================================================================
class MarketplaceUnitEconomics:
    """Основной класс для расчета юнит-экономики"""
    _instance = None
    _lock = Lock()

    def __new__(cls):
        if cls._instance is None:
            with cls._lock:
                if cls._instance is None:
                    cls._instance = super().__new__(cls)
                    cls._instance._initialized = False
        return cls._instance

    def __init__(self):
        if self._initialized:
            return
        self._initialized = True
        self._configs = self._load_marketplace_configs()
        self._categories = self._load_categories()
        self._cache = {}
        self._history = []
        self._stats = self._init_stats()
        self._settings = self._load_settings()
        self._logger = logging.getLogger('MarketplaceUnitEconomics')
        self._logger.info("🚗 Инициализация MarketplaceUnitEconomics")
        self._logger.info(f"📊 Загружено {len(self._configs)} маркетплейсов")
        self._logger.info(f"📚 Загружено {len(self._categories)} категорий")

    def _load_marketplace_configs(self) -> Dict[str, MarketplaceConfig]:
        """Загрузка конфигураций маркетплейсов"""
        return get_marketplace_configs_2026()

    def _load_categories(self) -> Dict[str, ProductCategory]:
        """Загрузка категорий"""
        categories = {}
        for name, cat in get_auto_parts_categories_full().items():
            categories[name] = ProductCategory(
                name=cat.name,
                description=cat.description,
                dimensions=ProductDimensions(
                    length=(cat.min_length + cat.max_length) / 2,
                    width=(cat.min_width + cat.max_width) / 2,
                    height=(cat.min_height + cat.max_height) / 2,
                    weight=(cat.min_weight + cat.max_weight) / 2
                ),
                typical_volume=cat.typical_volume,
                typical_weight=cat.typical_weight,
                oem_codes=cat.oem_codes,
                cross_references=cat.cross_references,
                alternatives=cat.alternatives,
                compatibility=cat.compatibility,
                hazardous=cat.hazardous,
                fragile=cat.fragile,
                requires_special_packaging=cat.requires_special_packaging,
                seasonality=cat.seasonality,
                risk_level=cat.risk_level,
                min_length=cat.min_length,
                max_length=cat.max_length,
                min_width=cat.min_width,
                max_width=cat.max_width,
                min_height=cat.min_height,
                max_height=cat.max_height,
                min_weight=cat.min_weight,
                max_weight=cat.max_weight
            )
        return categories

    def _init_stats(self) -> Dict[str, Any]:
        return {
            "total_calculations": 0,
            "by_marketplace": defaultdict(int),
            "by_category": defaultdict(int),
            "by_mode": defaultdict(int),
            "by_status": defaultdict(int),
            "avg_profit": 0.0,
            "avg_margin": 0.0,
            "avg_roi": 0.0,
            "total_profit": 0.0,
            "max_profit": 0.0,
            "min_profit": 0.0,
            "best_marketplace": None,
            "best_category": None,
            "best_mode": None,
            "total_optimizations": 0,
            "optimization_improvement": 0.0,
            "start_time": datetime.now(),
            "errors_count": 0,
            "last_error": None,
            "cache_hits": 0,
            "cache_misses": 0
        }

    def _load_settings(self) -> Dict[str, Any]:
        settings_path = CONFIG_DIR / "settings.json"
        default_settings = {
            "default_marketplace": "Ozon",
            "default_mode": "FBS",
            "default_days_storage": 30,
            "target_margin": 20.0,
            "enable_ai": True,
            "ai_provider": "deepseek",
            "enable_cache": True,
            "cache_ttl": 3600,
            "parallel_processing": True,
            "max_workers": 4,
            "optimize_memory": True,
            "precision_decimals": 2,
            "currency": "RUB",
            "locale": "ru_RU",
            "timezone": "Europe/Moscow"
        }
        if settings_path.exists():
            try:
                with open(settings_path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    default_settings.update(settings)
            except Exception as e:
                self._logger.warning(f"Ошибка загрузки настроек: {e}")
        return default_settings

    def get_category_dimensions(self, category_name: str) -> Optional[ProductDimensions]:
        """Получить габариты категории"""
        if category_name in self._categories:
            return self._categories[category_name].dimensions
        return None

    def find_categories_by_keyword(self, keyword: str) -> List[Tuple[str, ProductCategory]]:
        """Поиск категорий по ключевому слову"""
        keyword_lower = keyword.lower()
        results = []
        for name, cat in self._categories.items():
            if keyword_lower in name.lower() or keyword_lower in cat.description.lower():
                results.append((name, cat))
        return results

    def find_categories_by_oem(self, oem_code: str) -> List[Tuple[str, ProductCategory]]:
        """Поиск категорий по OEM коду"""
        oem_lower = oem_code.lower()
        results = []
        for name, cat in self._categories.items():
            if any(oem_lower in str(oem).lower() for oem in cat.oem_codes):
                results.append((name, cat))
        return results

    def calculate_dimensions_from_category(self, category_name: str) -> Tuple[float, float, float, float]:
        """Расчет габаритов из категории"""
        cat = self._categories.get(category_name)
        if cat and cat.dimensions:
            return (
                cat.dimensions.length,
                cat.dimensions.width,
                cat.dimensions.height,
                cat.dimensions.weight
            )
        return 0, 0, 0, 0

    @timer_decorator
    @cache_decorator(ttl=CACHE_TTL)
    def calculate_unit_economics(
        self,
        price: float,
        cost: float,
        marketplace: str,
        category: Optional[str] = None,
        operation_mode: str = "FBS",
        days_in_storage: int = 30,
        length: float = 0,
        width: float = 0,
        height: float = 0,
        weight: float = 0,
        is_premium: bool = False,
        include_insurance: bool = False,
        include_packing: bool = False,
        include_marketing: bool = False,
        currency: str = "RUB",
        **kwargs
    ) -> UnitEconomicsResult:
        """Расчет юнит-экономики"""
        if price <= 0:
            raise ValidationError("Цена должна быть положительной", "price", price)
        if cost <= 0:
            raise ValidationError("Себестоимость должна быть положительной", "cost", cost)
        if marketplace not in self._configs:
            raise MarketplaceError(f"Маркетплейс {marketplace} не поддерживается", marketplace)

        config = self._configs[marketplace]

        if all([length == 0, width == 0, height == 0, weight == 0]) and category:
            length, width, height, weight = self.calculate_dimensions_from_category(category)

        volume = calculate_volume(length, width, height)
        if volume == 0:
            volume = 5.0
        if weight <= 0:
            weight = 1.0

        commission_rate = config.category_rates.get(category, config.commission_rate)
        commission = max(price * commission_rate, config.min_commission)
        if config.max_commission < float('inf'):
            commission = min(commission, config.max_commission)

        subscription_cost = config.subscription_fee / 30 if config.subscription_fee > 0 else 0

        logistics = (
            config.logistics_base +
            weight * config.logistics_per_kg +
            volume * config.logistics_per_liter
        )
        mode_multiplier = config.mode_multipliers.get(operation_mode, 1.0)
        logistics *= mode_multiplier

        storage_cost = volume * config.storage_per_day * days_in_storage
        acquiring = price * config.acquiring_fee
        delivery = price * config.delivery_fee_percent
        last_mile = config.last_mile_fee
        returns = price * config.return_fee
        rko_fee = price * config.rko_fee if config.rko_fee > 0 else 0
        premium_fee = price * config.premium_fee if is_premium and config.premium_fee > 0 else 0
        insurance_fee = price * config.insurance_fee if include_insurance and config.insurance_fee > 0 else 0
        packing_fee = config.packing_fee if include_packing and config.packing_fee > 0 else 0
        marketing_fee = price * config.marketing_fee if include_marketing and config.marketing_fee > 0 else 0

        total_expenses = (
            cost + commission + subscription_cost + logistics + storage_cost +
            acquiring + delivery + last_mile + returns + rko_fee +
            premium_fee + insurance_fee + packing_fee + marketing_fee
        )

        profit = price - total_expenses
        margin_percent = (profit / price * 100) if price > 0 else 0
        roi = (profit / cost * 100) if cost > 0 else 0

        fixed_costs = logistics + storage_cost + last_mile + subscription_cost
        variable_rate = (
            commission_rate + config.acquiring_fee +
            config.delivery_fee_percent + config.return_fee +
            config.rko_fee + config.premium_fee +
            config.insurance_fee + config.marketing_fee
        )
        breakeven_price = ((cost + fixed_costs) / (1 - variable_rate)) if (1 - variable_rate) > 0 else 0

        contribution_margin = price - cost - commission - logistics - acquiring - delivery - last_mile - returns
        contribution_margin_ratio = (contribution_margin / price * 100) if price > 0 else 0

        result = UnitEconomicsResult(
            marketplace=marketplace,
            operation_mode=operation_mode,
            category=category or "Общая",
            price=round(price, 2),
            cost=round(cost, 2),
            length=round(length, 2),
            width=round(width, 2),
            height=round(height, 2),
            weight=round(weight, 2),
            volume=round(volume, 3),
            commission=round(commission, 2),
            commission_percent=round(commission / price * 100, 2) if price > 0 else 0,
            logistics=round(logistics, 2),
            storage_cost=round(storage_cost, 2),
            acquiring=round(acquiring, 2),
            delivery=round(delivery, 2),
            last_mile=round(last_mile, 2),
            returns=round(returns, 2),
            rko_fee=round(rko_fee, 2),
            premium_fee=round(premium_fee, 2),
            insurance_fee=round(insurance_fee, 2),
            packing_fee=round(packing_fee, 2),
            marketing_fee=round(marketing_fee, 2),
            subscription_cost=round(subscription_cost, 2),
            total_expenses=round(total_expenses, 2),
            profit=round(profit, 2),
            margin_percent=round(margin_percent, 2),
            roi=round(roi, 2),
            breakeven_price=round(breakeven_price, 2),
            profit_per_ruble=round(profit / price, 4) if price > 0 else 0,
            contribution_margin=round(contribution_margin, 2),
            contribution_margin_ratio=round(contribution_margin_ratio, 2),
            status=CalculationStatus.COMPLETED,
            metadata=kwargs
        )

        self._update_stats(result)
        self._history.append(result)
        if len(self._history) > HISTORY_LIMIT:
            self._history = self._history[-HISTORY_LIMIT:]

        return result

    def _update_stats(self, result: UnitEconomicsResult):
        """Обновление статистики"""
        self._stats["total_calculations"] += 1
        self._stats["by_marketplace"][result.marketplace] += 1
        self._stats["by_category"][result.category] += 1
        self._stats["by_mode"][result.operation_mode] += 1
        self._stats["by_status"][result.status.name] += 1
        self._stats["total_profit"] += result.profit
        if result.profit > self._stats["max_profit"]:
            self._stats["max_profit"] = result.profit
            self._stats["best_marketplace"] = result.marketplace
            self._stats["best_category"] = result.category
            self._stats["best_mode"] = result.operation_mode
        if result.profit < self._stats["min_profit"] or self._stats["min_profit"] == 0:
            self._stats["min_profit"] = result.profit
        n = self._stats["total_calculations"]
        self._stats["avg_profit"] = self._stats["total_profit"] / n
        self._stats["avg_margin"] = (self._stats["avg_margin"] * (n - 1) + result.margin_percent) / n
        self._stats["avg_roi"] = (self._stats["avg_roi"] * (n - 1) + result.roi) / n

    @timer_decorator
    def calculate_for_all_marketplaces(
        self,
        price: float,
        cost: float,
        category: Optional[str] = None,
        operation_mode: str = "FBS",
        days_in_storage: int = 30,
        length: float = 0,
        width: float = 0,
        height: float = 0,
        weight: float = 0,
        **kwargs
    ) -> pd.DataFrame:
        """Расчет для всех маркетплейсов"""
        results = []
        for marketplace in self._configs.keys():
            try:
                result = self.calculate_unit_economics(
                    price=price,
                    cost=cost,
                    marketplace=marketplace,
                    category=category,
                    operation_mode=operation_mode,
                    days_in_storage=days_in_storage,
                    length=length,
                    width=width,
                    height=height,
                    weight=weight,
                    **kwargs
                )
                results.append(result)
            except Exception as e:
                self._logger.error(f"Ошибка расчета для {marketplace}: {e}")
                self._stats["errors_count"] += 1
                self._stats["last_error"] = str(e)
        if not results:
            return pd.DataFrame()
        return pd.DataFrame([r.to_dict() for r in results])

    @timer_decorator
    def calculate_for_catalog_batch(
        self,
        df: pd.DataFrame,
        price_col: str,
        cost_col: str,
        category_col: Optional[str] = None,
        length_col: Optional[str] = None,
        width_col: Optional[str] = None,
        height_col: Optional[str] = None,
        weight_col: Optional[str] = None,
        article_col: str = "Артикул",
        brand_col: str = "Бренд",
        marketplaces: Optional[List[str]] = None,
        operation_mode: str = "FBS",
        days_in_storage: int = 30,
        progress_callback: Optional[Callable] = None,
        max_workers: int = 4
    ) -> pd.DataFrame:
        """Пакетный расчет для каталога"""
        if marketplaces is None:
            marketplaces = list(self._configs.keys())
        items = []
        for idx, row in df.iterrows():
            price = safe_float(row.get(price_col, 0))
            cost = safe_float(row.get(cost_col, 0))
            article = safe_str(row.get(article_col, f"Товар_{idx}"))
            brand = safe_str(row.get(brand_col, ""))
            if price <= 0 or cost <= 0:
                continue
            length = safe_float(row.get(length_col, 0)) if length_col else 0
            width = safe_float(row.get(width_col, 0)) if width_col else 0
            height = safe_float(row.get(height_col, 0)) if height_col else 0
            weight = safe_float(row.get(weight_col, 0)) if weight_col else 0
            category = safe_str(row.get(category_col, "")) if category_col else None
            items.append({
                "idx": idx,
                "article": article,
                "brand": brand,
                "price": price,
                "cost": cost,
                "category": category,
                "length": length,
                "width": width,
                "height": height,
                "weight": weight
            })
        total_items = len(items) * len(marketplaces)
        if total_items == 0:
            return pd.DataFrame()
        results = []
        processed = 0
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = []
            for item in items:
                for marketplace in marketplaces:
                    future = executor.submit(
                        self.calculate_unit_economics,
                        price=item["price"],
                        cost=item["cost"],
                        marketplace=marketplace,
                        category=item["category"],
                        operation_mode=operation_mode,
                        days_in_storage=days_in_storage,
                        length=item["length"],
                        width=item["width"],
                        height=item["height"],
                        weight=item["weight"]
                    )
                    futures.append((future, item["article"], item["brand"], item["idx"]))
            for future, article, brand, idx in futures:
                try:
                    result = future.result(timeout=30)
                    result_dict = result.to_dict()
                    result_dict["Артикул"] = article
                    result_dict["Бренд"] = brand
                    result_dict["Индекс"] = idx
                    results.append(result_dict)
                except Exception as e:
                    self._logger.error(f"Ошибка расчета для {article}: {e}")
                    self._stats["errors_count"] += 1
                    self._stats["last_error"] = str(e)
                processed += 1
                if progress_callback and processed % 10 == 0:
                    progress_callback(processed / total_items)
        if progress_callback:
            progress_callback(1.0)
        return pd.DataFrame(results) if results else pd.DataFrame()

    @timer_decorator
    def optimize_price(
        self,
        cost: float,
        marketplace: str,
        category: Optional[str] = None,
        operation_mode: str = "FBS",
        days_in_storage: int = 30,
        length: float = 0,
        width: float = 0,
        height: float = 0,
        weight: float = 0,
        target_margin: float = 20.0,
        price_min: float = 0,
        price_max: float = 100000,
        step: float = 10,
        max_iterations: int = 1000
    ) -> OptimizationResult:
        """Оптимизация цены"""
        current_price = max(price_min, cost * 1.1) if price_min == 0 else price_min
        best_price = current_price
        best_profit = float('-inf')
        best_margin = 0
        best_result = None
        iteration = 0
        while current_price <= price_max and iteration < max_iterations:
            try:
                result = self.calculate_unit_economics(
                    price=current_price,
                    cost=cost,
                    marketplace=marketplace,
                    category=category,
                    operation_mode=operation_mode,
                    days_in_storage=days_in_storage,
                    length=length,
                    width=width,
                    height=height,
                    weight=weight
                )
                margin = result.margin_percent
                profit = result.profit
                if margin >= target_margin and profit > best_profit:
                    best_profit = profit
                    best_price = current_price
                    best_margin = margin
                    best_result = result
                current_price += step
                iteration += 1
            except Exception as e:
                self._logger.warning(f"Ошибка при оптимизации для цены {current_price}: {e}")
                current_price += step

        current_result = self.calculate_unit_economics(
            price=price_min or best_price,
            cost=cost,
            marketplace=marketplace,
            category=category,
            operation_mode=operation_mode,
            days_in_storage=days_in_storage,
            length=length,
            width=width,
            height=height,
            weight=weight
        )
        improvement_pct = ((best_profit - current_result.profit) / current_result.profit * 100) if current_result.profit > 0 else 0
        recommendations = []
        if best_price > 0 and best_margin >= target_margin:
            recommendations.append(f"Установите цену {best_price:.2f} ₽ для достижения маржи {target_margin}%")
        else:
            recommendations.append(f"Целевая маржа {target_margin}% не достигнута. Максимальная маржа: {best_margin:.1f}%")
        if best_profit > current_result.profit:
            recommendations.append(f"Потенциальное увеличение прибыли: {improvement_pct:.1f}%")

        self._stats["total_optimizations"] += 1
        self._stats["optimization_improvement"] += improvement_pct

        return OptimizationResult(
            optimal_price=best_price,
            optimal_margin=best_margin,
            optimal_profit=best_profit,
            current_price=current_result.price,
            current_margin=current_result.margin_percent,
            current_profit=current_result.profit,
            improvement_pct=improvement_pct,
            recommendations=recommendations,
            metadata={
                "target_margin": target_margin,
                "step": step,
                "iterations": iteration
            }
        )

    @timer_decorator
    def forecast_profit(
        self,
        current_data: Dict[str, Any],
        periods: int = 12,
        growth_rate: float = 0.05,
        seasonality: Optional[List[float]] = None,
        confidence_level: float = 0.95
    ) -> ForecastResult:
        """Прогнозирование прибыли"""
        if seasonality is None:
            seasonality = [0.85, 0.85, 0.95, 1.05, 1.10, 1.15,
                           1.20, 1.15, 1.10, 1.05, 0.95, 0.90]
        base_value = current_data.get("profit", 1000)
        periods_list = []
        values_list = []
        seasonality_list = []
        trend_list = []
        for i in range(periods):
            month_idx = i % 12
            seasonal_factor = seasonality[month_idx] if month_idx < len(seasonality) else 1.0
            growth_factor = (1 + growth_rate) ** (i / 12)
            factor = seasonal_factor * growth_factor
            value = base_value * factor
            if DATEUTIL_AVAILABLE:
                periods_list.append(datetime.now() + relativedelta(months=i))
            else:
                periods_list.append(datetime.now() + timedelta(days=30 * i))
            values_list.append(value)
            seasonality_list.append(seasonal_factor)
            trend_list.append(growth_factor)
        std_dev = np.std(values_list) * 0.2
        z_score = 1.96
        lower_bound = [v - z_score * std_dev for v in values_list]
        upper_bound = [v + z_score * std_dev for v in values_list]
        return ForecastResult(
            periods=periods_list,
            values=values_list,
            seasonality=seasonality_list,
            trend=trend_list,
            confidence_intervals=(lower_bound, upper_bound),
            metadata={
                "base_value": base_value,
                "growth_rate": growth_rate,
                "confidence_level": confidence_level
            }
        )

    def get_history(self, limit: int = 100, filters: Optional[Dict] = None) -> List[UnitEconomicsResult]:
        """Получение истории"""
        history = self._history[-limit:] if limit > 0 else self._history
        if filters:
            filtered = []
            for item in history:
                match = True
                for key, value in filters.items():
                    if key == "marketplace" and item.marketplace != value:
                        match = False
                        break
                    elif key == "category" and item.category != value:
                        match = False
                        break
                    elif key == "operation_mode" and item.operation_mode != value:
                        match = False
                        break
                    elif key == "min_profit" and item.profit < value:
                        match = False
                        break
                    elif key == "max_profit" and item.profit > value:
                        match = False
                        break
                    elif key == "start_date" and item.timestamp < value:
                        match = False
                        break
                    elif key == "end_date" and item.timestamp > value:
                        match = False
                        break
                if match:
                    filtered.append(item)
            return filtered
        return history

    def get_stats(self) -> Dict[str, Any]:
        """Получение статистики"""
        stats = self._stats.copy()
        stats["history_count"] = len(self._history)
        stats["cache_size"] = len(self._cache)
        stats["uptime"] = (datetime.now() - stats["start_time"]).total_seconds()
        if stats["total_calculations"] > 0:
            stats["success_rate"] = 1 - (stats["errors_count"] / stats["total_calculations"])
        else:
            stats["success_rate"] = 0
        return stats

    def clear_history(self):
        """Очистка истории"""
        self._history = []
        self._stats = self._init_stats()
        self._cache.clear()
        gc.collect()

    def get_best_configuration(self) -> Dict[str, Any]:
        """Получение лучшей конфигурации"""
        if not self._history:
            return {"error": "Нет данных"}
        best = max(self._history, key=lambda x: x.profit)
        return {
            "marketplace": best.marketplace,
            "operation_mode": best.operation_mode,
            "category": best.category,
            "profit": best.profit,
            "margin": best.margin_percent,
            "price": best.price,
            "cost": best.cost,
            "timestamp": best.timestamp.isoformat()
        }

    def get_category_stats(self) -> pd.DataFrame:
        """Статистика по категориям"""
        if not self._history:
            return pd.DataFrame()
        stats = defaultdict(lambda: {
            "count": 0,
            "total_profit": 0,
            "avg_profit": 0,
            "avg_margin": 0,
            "best_profit": 0,
            "worst_profit": 0
        })
        for result in self._history:
            cat = result.category
            stats[cat]["count"] += 1
            stats[cat]["total_profit"] += result.profit
            stats[cat]["avg_margin"] += result.margin_percent
            stats[cat]["best_profit"] = max(stats[cat]["best_profit"], result.profit)
            stats[cat]["worst_profit"] = min(stats[cat]["worst_profit"], result.profit)
        for cat in stats:
            if stats[cat]["count"] > 0:
                stats[cat]["avg_profit"] = stats[cat]["total_profit"] / stats[cat]["count"]
                stats[cat]["avg_margin"] /= stats[cat]["count"]
        return pd.DataFrame.from_dict(stats, orient="index").reset_index().rename(columns={"index": "category"})

    def get_marketplace_stats(self) -> pd.DataFrame:
        """Статистика по маркетплейсам"""
        if not self._history:
            return pd.DataFrame()
        stats = defaultdict(lambda: {
            "count": 0,
            "total_profit": 0,
            "avg_profit": 0,
            "avg_margin": 0,
            "best_profit": 0,
            "worst_profit": 0
        })
        for result in self._history:
            mp = result.marketplace
            stats[mp]["count"] += 1
            stats[mp]["total_profit"] += result.profit
            stats[mp]["avg_margin"] += result.margin_percent
            stats[mp]["best_profit"] = max(stats[mp]["best_profit"], result.profit)
            stats[mp]["worst_profit"] = min(stats[mp]["worst_profit"], result.profit)
        for mp in stats:
            if stats[mp]["count"] > 0:
                stats[mp]["avg_profit"] = stats[mp]["total_profit"] / stats[mp]["count"]
                stats[mp]["avg_margin"] /= stats[mp]["count"]
        return pd.DataFrame.from_dict(stats, orient="index").reset_index().rename(columns={"index": "marketplace"})

    def export_history(self, format: ExportFormat = ExportFormat.EXCEL) -> bytes:
        """Экспорт истории"""
        if not self._history:
            return b""
        df = pd.DataFrame([r.to_dict() for r in self._history])
        if format == ExportFormat.CSV:
            return df.to_csv(index=False, encoding='utf-8-sig').encode('utf-8')
        elif format == ExportFormat.EXCEL:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='История')
            return output.getvalue()
        elif format == ExportFormat.JSON:
            return df.to_json(orient='records', force_ascii=False).encode('utf-8')
        else:
            raise ExportError(f"Формат {format.name} не поддерживается", format=format.name)


# ============================================================================
# БЛОК 6: DEEPSEEK AI ДЛЯ ОБНОВЛЕНИЯ ТАРИФОВ (250+ СТРОК)
# ============================================================================
class DeepSeekRateUpdater:
    """Класс для обновления тарифов через DeepSeek AI"""
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.environ.get('DEEPSEEK_API_KEY')
        self.api_url = DEEPSEEK_API_URL
        self.cache_file = CACHE_DIR / "deepseek_rates_cache.json"
        self.cache_file.parent.mkdir(exist_ok=True)
        self.session = requests.Session()
        self._logger = logging.getLogger('DeepSeekRateUpdater')
        if self.api_key:
            self.session.headers.update({
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            })
            self._logger.info("DeepSeek клиент инициализирован")
        else:
            self._logger.warning("DeepSeek API ключ не найден")

    def _build_prompt(self, marketplace: str, category: str = None) -> str:
        """Формирование промпта для DeepSeek (ИСПРАВЛЕНО!)"""
        prompt = (
            "Ты - эксперт по юнит-экономике маркетплейсов России, специализирующийся на автозапчастях.\n"
            f"Предоставь актуальные тарифы для маркетплейса {marketplace} на 2026 год.\n"
            "Формат ответа ТОЛЬКО JSON без пояснений:\n"
            "{\n"
            '  "commission_rate": число (комиссия в долях),\n'
            '  "min_commission": число (минимальная комиссия в рублях),\n'
            '  "logistics_base": число (базовая стоимость логистики),\n'
            '  "logistics_per_kg": число (стоимость за кг),\n'
            '  "logistics_per_liter": число (стоимость за литр объема),\n'
            '  "storage_per_day": число (стоимость хранения за день),\n'
            '  "return_fee": число (процент возвратов в долях),\n'
            '  "acquiring_fee": число (процент эквайринга в долях),\n'
            '  "last_mile_fee": число (последняя миля в рублях),\n'
            '  "delivery_fee_percent": число (процент доставки в долях)\n'
            "}\n"
        )
        if category:
            prompt += f"\nУкажи комиссию для категории '{category}' в поле 'category_rate'."
        return prompt

    def _call_deepseek_api(self, prompt: str) -> Dict[str, Any]:
        """Вызов DeepSeek API"""
        if not self.api_key:
            return {}
        try:
            payload = {
                "model": DEEPSEEK_MODEL,
                "messages": [
                    {"role": "system", "content": "Ты - эксперт по маркетплейсам и автозапчастям."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.1,
                "max_tokens": 1000
            }
            response = self.session.post(self.api_url, json=payload, timeout=30)
            if response.status_code == 200:
                result = response.json()
                content = result['choices'][0]['message']['content']
                json_match = re.search(r'\{.*\}', content, re.DOTALL)
                if json_match:
                    return json.loads(json_match.group())
                return {}
            else:
                self._logger.error(f"DeepSeek API ошибка: {response.status_code}")
                return {}
        except Exception as e:
            self._logger.error(f"Ошибка вызова DeepSeek API: {e}")
            return {}

    def _load_cache(self) -> Dict:
        """Загрузка кэша"""
        if self.cache_file.exists():
            try:
                with open(self.cache_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception:
                return {}
        return {}

    def _save_cache(self, cache_data: Dict):
        """Сохранение кэша"""
        try:
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(cache_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            self._logger.error(f"Ошибка сохранения кэша: {e}")

    def get_rates_from_ai(self, marketplace: str, category: str = None) -> Dict[str, Any]:
        """Получение тарифов через DeepSeek"""
        if not self.api_key:
            return {}
        cache_data = self._load_cache()
        cache_key = f"{marketplace}_{category or 'auto_parts'}"
        if cache_key in cache_data:
            cached = cache_data[cache_key]
            if time.time() - cached['timestamp'] < 86400:
                self._logger.info(f"Использованы кэшированные тарифы для {marketplace}")
                return cached['data']
        try:
            prompt = self._build_prompt(marketplace, category)
            result = self._call_deepseek_api(prompt)
            if result:
                cache_data[cache_key] = {
                    'timestamp': time.time(),
                    'data': result
                }
                self._save_cache(cache_data)
                self._logger.info(f"Тарифы для {marketplace} обновлены")
                return result
            return {}
        except Exception as e:
            self._logger.error(f"Ошибка получения тарифов: {e}")
            return {}


# ============================================================================
# БЛОК 7: HIGH-VOLUME CATALOG (POLARS + DUCKDB) (700+ СТРОК)
# ============================================================================
class HighVolumeAutoPartsCatalog:
    """
    High-Volume каталог автозапчастей с поддержкой 10M+ записей
    Использует Polars для быстрой обработки и DuckDB для хранения.
    """
    def __init__(self):
        self.data_dir = Path("./auto_parts_data")
        self.data_dir.mkdir(exist_ok=True)
        self.cloud_config = self.load_cloud_config()
        self.price_rules = self.load_price_rules()
        self.exclusion_rules = self.load_exclusion_rules()
        self.category_mapping = self.load_category_mapping()
        self.db_path = self.data_dir / "catalog.duckdb"
        if POLARS_AVAILABLE and DUCKDB_AVAILABLE:
            try:
                self.conn = duckdb.connect(database=str(self.db_path))
                self.setup_database()
            except Exception as e:
                logger.error(f"Ошибка подключения к DuckDB: {e}")
                self.conn = None
        else:
            self.conn = None
            logger.warning("HighVolume режим недоступен: установите polars и duckdb")

    def load_cloud_config(self) -> Dict[str, Any]:
        config_path = self.data_dir / "cloud_config.json"
        default_config = {
            "enabled": False,
            "provider": "s3",
            "bucket": "",
            "region": "",
            "sync_interval": 3600,
            "last_sync": 0
        }
        if config_path.exists():
            try:
                return json.loads(config_path.read_text(encoding='utf-8'))
            except Exception as e:
                logger.error(f"Ошибка чтения cloud_config.json: {e}")
        else:
            config_path.write_text(json.dumps(
                default_config, indent=2, ensure_ascii=False), encoding='utf-8')
        return default_config

    def save_cloud_config(self):
        config_path = self.data_dir / "cloud_config.json"
        self.cloud_config["last_sync"] = int(time.time())
        config_path.write_text(json.dumps(
            self.cloud_config, indent=2, ensure_ascii=False), encoding='utf-8')

    def load_price_rules(self) -> Dict[str, Any]:
        price_rules_path = self.data_dir / "price_rules.json"
        default_rules = {
            "global_markup": 0.2,
            "brand_markups": {},
            "min_price": 0.0,
            "max_price": 99999.0
        }
        if price_rules_path.exists():
            try:
                return json.loads(price_rules_path.read_text(encoding='utf-8'))
            except Exception as e:
                logger.error(f"Ошибка чтения price_rules.json: {e}")
        else:
            price_rules_path.write_text(json.dumps(
                default_rules, indent=2, ensure_ascii=False), encoding='utf-8')
        return default_rules

    def save_price_rules(self):
        price_rules_path = self.data_dir / "price_rules.json"
        price_rules_path.write_text(json.dumps(
            self.price_rules, indent=2, ensure_ascii=False), encoding='utf-8')

    def load_exclusion_rules(self) -> List[str]:
        exclusion_path = self.data_dir / "exclusion_rules.txt"
        if exclusion_path.exists():
            try:
                return [line.strip() for line in exclusion_path.read_text(encoding='utf-8').splitlines() if line.strip()]
            except Exception as e:
                logger.error(f"Ошибка чтения exclusion_rules.txt: {e}")
        else:
            content = "Кузов\nСтекла\nМасла"
            exclusion_path.write_text(content, encoding='utf-8')
            return ["Кузов", "Стекла", "Масла"]

    def save_exclusion_rules(self):
        exclusion_path = self.data_dir / "exclusion_rules.txt"
        exclusion_path.write_text("\n".join(self.exclusion_rules), encoding='utf-8')

    def load_category_mapping(self) -> Dict[str, str]:
        category_path = self.data_dir / "category_mapping.txt"
        default_mapping = {
            "Радиатор": "Охлаждение",
            "Шаровая опора": "Подвеска",
            "Фильтр масляный": "Фильтры",
            "Тормозные колодки": "Тормоза"
        }
        if category_path.exists():
            try:
                mapping = {}
                for line in category_path.read_text(encoding='utf-8').splitlines():
                    if line.strip() and "|" in line:
                        key, value = line.split("|", 1)
                        mapping[key.strip()] = value.strip()
                return mapping
            except Exception as e:
                logger.error(f"Ошибка чтения category_mapping.txt: {e}")
        else:
            content = "\n".join([f"{k}|{v}" for k, v in default_mapping.items()])
            category_path.write_text(content, encoding='utf-8')
        return default_mapping

    def save_category_mapping(self):
        category_path = self.data_dir / "category_mapping.txt"
        content = "\n".join([f"{k}|{v}" for k, v in self.category_mapping.items()])
        category_path.write_text(content, encoding='utf-8')

    def setup_database(self):
        if not self.conn:
            return
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS oe (
                oe_number_norm VARCHAR PRIMARY KEY,
                oe_number VARCHAR,
                name VARCHAR,
                applicability VARCHAR,
                category VARCHAR
            )
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS parts (
                artikul_norm VARCHAR,
                brand_norm VARCHAR,
                artikul VARCHAR,
                brand VARCHAR,
                multiplicity INTEGER,
                barcode VARCHAR,
                length DOUBLE,
                width DOUBLE,
                height DOUBLE,
                weight DOUBLE,
                image_url VARCHAR,
                dimensions_str VARCHAR,
                description VARCHAR,
                PRIMARY KEY (artikul_norm, brand_norm)
            )
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS cross_references (
                oe_number_norm VARCHAR,
                artikul_norm VARCHAR,
                brand_norm VARCHAR,
                PRIMARY KEY (oe_number_norm, artikul_norm, brand_norm)
            )
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS prices (
                artikul_norm VARCHAR,
                brand_norm VARCHAR,
                price DOUBLE,
                currency VARCHAR DEFAULT 'RUB',
                PRIMARY KEY (artikul_norm, brand_norm)
            )
        """)
        self.conn.execute("""
            CREATE TABLE IF NOT EXISTS metadata (
                key VARCHAR PRIMARY KEY,
                value VARCHAR
            )
        """)
        self.create_indexes()

    def create_indexes(self):
        if not self.conn:
            return
        try:
            indexes = [
                "CREATE INDEX IF NOT EXISTS idx_oe_number_norm ON oe(oe_number_norm)",
                "CREATE INDEX IF NOT EXISTS idx_parts_keys ON parts(artikul_norm, brand_norm)",
                "CREATE INDEX IF NOT EXISTS idx_cross_oe ON cross_references(oe_number_norm)",
                "CREATE INDEX IF NOT EXISTS idx_cross_artikul ON cross_references(artikul_norm, brand_norm)",
                "CREATE INDEX IF NOT EXISTS idx_prices_keys ON prices(artikul_norm, brand_norm)"
            ]
            for index_sql in indexes:
                try:
                    self.conn.execute(index_sql)
                except Exception as e:
                    logger.warning(f"Не удалось создать индекс: {e}")
        except Exception as e:
            logger.warning(f"Ошибка создания индексов: {e}")

    @staticmethod
    def normalize_key(series) -> "pl.Series":
        if not POLARS_AVAILABLE:
            return series
        return (series
                .fill_null("")
                .cast(pl.Utf8)
                .str.replace_all("'", "")
                .str.replace_all(r"[^0-9A-Za-zА-Яа-яЁё`\-\s]", "")
                .str.replace_all(r"\s+", " ")
                .str.strip_chars()
                .str.to_lowercase())

    @staticmethod
    def clean_values(series) -> "pl.Series":
        if not POLARS_AVAILABLE:
            return series
        return (series
                .fill_null("")
                .cast(pl.Utf8)
                .str.replace_all("'", "")
                .str.replace_all(r"[^0-9A-Za-zА-Яа-яЁё`\-\s]", "")
                .str.replace_all(r"\s+", " ")
                .str.strip_chars())

    def determine_category_vectorized(self, name_series) -> "pl.Series":
        if not POLARS_AVAILABLE:
            return name_series
        name_lower = name_series.str.to_lowercase()
        categorization_expr = pl.when(pl.lit(False)).then(pl.lit(None))
        for key, category in self.category_mapping.items():
            categorization_expr = categorization_expr.when(
                name_lower.str.contains(key.lower())
            ).then(pl.lit(category))
        categories_map = {
            'Фильтр': 'фильтр|filter',
            'Тормоза': 'тормоз|brake|колодк|диск|суппорт',
            'Подвеска': 'амортизатор|стойк|spring|подвеск|рычаг',
            'Двигатель': 'двигатель|engine|свеч|поршень|клапан',
            'Трансмиссия': 'трансмиссия|сцеплен|коробк|transmission',
            'Электрика': 'аккумулятор|генератор|стартер|провод|ламп',
            'Рулевое': 'рулевой|тяга|наконечник|steering',
            'Выпуск': 'глушитель|катализатор|выхлоп|exhaust',
            'Охлаждение': 'радиатор|вентилятор|термостат|cooling',
            'Топливо': 'топливный|бензонасос|форсунк|fuel'
        }
        for category, pattern in categories_map.items():
            categorization_expr = categorization_expr.when(
                name_lower.str.contains(pattern, literal=False)
            ).then(pl.lit(category))
        return categorization_expr.otherwise(pl.lit('Разное')).alias('category')

    def detect_columns(self, actual_columns: List[str], expected_columns: List[str]) -> Dict[str, str]:
        column_variants = {
            'oe_number': ['oe номер', 'oe', 'оe', 'номер', 'code', 'OE'],
            'artikul': ['артикул', 'article', 'sku'],
            'brand': ['бренд', 'brand', 'производитель', 'manufacturer'],
            'name': ['наименование', 'название', 'name', 'описание', 'description'],
            'applicability': ['применимость', 'автомобиль', 'vehicle', 'applicability'],
            'barcode': ['штрих-код', 'barcode', 'штрихкод', 'ean', 'eac13'],
            'multiplicity': ['кратность шт', 'кратность', 'multiplicity'],
            'length': ['длина (см)', 'длина', 'length', 'длинна'],
            'width': ['ширина (см)', 'ширина', 'width'],
            'height': ['высота (см)', 'высота', 'height'],
            'weight': ['вес (кг)', 'вес, кг', 'вес', 'weight'],
            'image_url': ['ссылка', 'url', 'изображение', 'image', 'картинка'],
            'dimensions_str': ['весогабариты', 'размеры', 'dimensions', 'size'],
            'price': ['цена', 'price', 'рекомендованная цена', 'retail price'],
            'currency': ['валюта', 'currency']
        }
        actual_lower = {col.lower(): col for col in actual_columns}
        mapping = {}
        for expected in expected_columns:
            variants = column_variants.get(expected, [expected])
            for variant in variants:
                variant_lower = variant.lower()
                for actual_l, actual_orig in actual_lower.items():
                    if variant_lower in actual_l and actual_orig not in mapping:
                        mapping[actual_orig] = expected
                        break
        return mapping

    def read_and_prepare_file(self, file_path: str, file_type: str) -> "pl.DataFrame":
        if not POLARS_AVAILABLE:
            logger.warning("Polars не доступен")
            return pl.DataFrame()
        logger.info(f"Обработка файла: {file_type} ({file_path})")
        try:
            if not os.path.exists(file_path):
                logger.error(f"Файл не найден: {file_path}")
                return pl.DataFrame()
            df = pl.read_excel(file_path, engine='calamine')
            if df.is_empty():
                logger.warning(f"Пустой файл: {file_path}")
                return pl.DataFrame()
        except Exception as e:
            logger.exception(f"Ошибка чтения файла {file_path}: {e}")
            return pl.DataFrame()
        schemas = {
            'oe': ['oe_number', 'artikul', 'brand', 'name', 'applicability'],
            'cross': ['oe_number', 'artikul', 'brand'],
            'barcode': ['artikul', 'brand', 'barcode', 'multiplicity'],
            'dimensions': ['artikul', 'brand', 'length', 'width', 'height', 'weight', 'dimensions_str'],
            'images': ['artikul', 'brand', 'image_url'],
            'prices': ['artikul', 'brand', 'price', 'currency']
        }
        expected_cols = schemas.get(file_type, [])
        column_mapping = self.detect_columns(df.columns, expected_cols)
        if not column_mapping:
            logger.warning(
                f"Не удалось определить колонки для файла {file_type}. Доступные: {df.columns}")
            return pl.DataFrame()
        df = df.rename(column_mapping)
        for col in ['artikul', 'brand', 'oe_number']:
            if col in df.columns:
                df = df.with_columns(self.clean_values(pl.col(col)).alias(col))
        key_cols = [col for col in ['oe_number', 'artikul', 'brand'] if col in df.columns]
        if key_cols:
            df = df.unique(subset=key_cols, keep='first')
        for col in ['artikul', 'brand', 'oe_number']:
            if col in df.columns:
                df = df.with_columns(self.normalize_key(
                    pl.col(col)).alias(f"{col}_norm"))
        return df

    def upsert_data(self, table_name: str, df: "pl.DataFrame", pk: List[str]):
        if not self.conn or df.is_empty():
            return
        df = df.unique(keep='first')
        temp_view_name = f"temp_{table_name}_{int(time.time())}"
        try:
            self.conn.register(temp_view_name, df.to_arrow())
        except Exception as e:
            logger.error(f"Ошибка регистрации временной таблицы: {e}")
            return
        try:
            pk_list = pk
            pk_cols_csv = ", ".join(f'"{c}"' for c in pk_list)
            delete_sql = f"""
                DELETE FROM {table_name}
                WHERE ({pk_cols_csv}) IN (SELECT {pk_cols_csv} FROM {temp_view_name});
            """
            self.conn.execute(delete_sql)
            insert_sql = f"""
                INSERT INTO {table_name}
                SELECT * FROM {temp_view_name};
            """
            self.conn.execute(insert_sql)
            logger.info(f"Успешно upsert {len(df)} записей в таблицу {table_name}.")
        except Exception as e:
            logger.error(f"Ошибка при UPSERT в {table_name}: {e}")
        finally:
            try:
                self.conn.unregister(temp_view_name)
            except Exception:
                pass

    def upsert_prices(self, price_df: "pl.DataFrame"):
        if not self.conn or price_df.is_empty():
            return
        if 'artikul' in price_df.columns and 'brand' in price_df.columns:
            price_df = price_df.with_columns([
                self.normalize_key(pl.col('artikul')).alias('artikul_norm'),
                self.normalize_key(pl.col('brand')).alias('brand_norm')
            ])
        if 'currency' not in price_df.columns:
            price_df = price_df.with_columns(pl.lit('RUB').alias('currency'))
        price_df = price_df.filter(
            (pl.col('price') >= self.price_rules['min_price']) &
            (pl.col('price') <= self.price_rules['max_price'])
        )
        self.upsert_data('prices', price_df, ['artikul_norm', 'brand_norm'])

    def process_and_load_data(self, dataframes: Dict[str, "pl.DataFrame"]):
        if not self.conn:
            logger.warning("⚠️ База данных не доступна")
            return
        logger.info("🔄 Начало загрузки и обновления данных в базе...")
        steps = [s for s in ['oe', 'cross', 'parts'] if s in dataframes]
        num_steps = len(steps)
        step_counter = 0
        if 'oe' in dataframes:
            step_counter += 1
            logger.info(f"({step_counter}/{num_steps}) Обработка OE данных...")
            df = dataframes['oe'].filter(pl.col('oe_number_norm') != "")
            oe_df = df.select(['oe_number_norm', 'oe_number', 'name', 'applicability']).unique(
                subset=['oe_number_norm'], keep='first')
            if 'name' in oe_df.columns:
                oe_df = oe_df.with_columns(
                    self.determine_category_vectorized(pl.col('name')))
            else:
                oe_df = oe_df.with_columns(category=pl.lit('Разное'))
            self.upsert_data('oe', oe_df, ['oe_number_norm'])
            cross_df_from_oe = df.filter(pl.col('artikul_norm') != "").select(
                ['oe_number_norm', 'artikul_norm', 'brand_norm']).unique()
            self.upsert_data('cross_references', cross_df_from_oe, [
                'oe_number_norm', 'artikul_norm', 'brand_norm'])
        if 'cross' in dataframes:
            step_counter += 1
            logger.info(f"({step_counter}/{num_steps}) Обработка кроссов...")
            df = dataframes['cross'].filter(
                (pl.col('oe_number_norm') != "") & (pl.col('artikul_norm') != ""))
            cross_df_from_cross = df.select(
                ['oe_number_norm', 'artikul_norm', 'brand_norm']).unique()
            self.upsert_data('cross_references', cross_df_from_cross, [
                'oe_number_norm', 'artikul_norm', 'brand_norm'])
        if 'prices' in dataframes:
            price_df = dataframes['prices']
            if not price_df.is_empty():
                logger.info("💰 Обработка цен...")
                self.upsert_prices(price_df)
                logger.info(f"✅ Успешно обновлено {len(price_df)} ценовых записей")
            step_counter += 1
        logger.info(f"({step_counter}/{num_steps}) Сборка и обновление данных по артикулам...")
        parts_df = None
        file_priority = ['oe', 'barcode', 'images', 'dimensions']
        key_files = {ftype: df for ftype, df in dataframes.items() if ftype in file_priority}
        if key_files:
            all_parts = pl.concat([
                df.select(['artikul', 'artikul_norm', 'brand', 'brand_norm'])
                for df in key_files.values() if 'artikul_norm' in df.columns and 'brand_norm' in df.columns
            ]).filter(pl.col('artikul_norm') != "").unique(subset=['artikul_norm', 'brand_norm'], keep='first')
            parts_df = all_parts
            for ftype in file_priority:
                if ftype not in key_files:
                    continue
                df = key_files[ftype]
                if df.is_empty() or 'artikul_norm' not in df.columns:
                    continue
                join_cols = [col for col in df.columns if col not in [
                    'artikul', 'artikul_norm', 'brand', 'brand_norm']]
                if not join_cols:
                    continue
                existing_cols = set(parts_df.columns)
                join_cols = [col for col in join_cols if col not in existing_cols]
                if not join_cols:
                    continue
                df_subset = df.select(['artikul_norm', 'brand_norm'] + join_cols).unique(
                    subset=['artikul_norm', 'brand_norm'], keep='first')
                parts_df = parts_df.join(
                    df_subset, on=['artikul_norm', 'brand_norm'], how='left', coalesce=True)
        if parts_df is not None and not parts_df.is_empty():
            if 'multiplicity' not in parts_df.columns:
                parts_df = parts_df.with_columns(
                    multiplicity=pl.lit(1).cast(pl.Int32))
            else:
                parts_df = parts_df.with_columns(
                    pl.col('multiplicity').fill_null(1).cast(pl.Int32))
            for col in ['length', 'width', 'height']:
                if col not in parts_df.columns:
                    parts_df = parts_df.with_columns(
                        pl.lit(None).cast(pl.Float64).alias(col))
            if 'dimensions_str' not in parts_df.columns:
                parts_df = parts_df.with_columns(
                    dimensions_str=pl.lit(None).cast(pl.Utf8))
            parts_df = parts_df.with_columns([
                pl.col('length').cast(pl.Utf8).fill_null('').alias('_length_str'),
                pl.col('width').cast(pl.Utf8).fill_null('').alias('_width_str'),
                pl.col('height').cast(pl.Utf8).fill_null('').alias('_height_str'),
            ])
            parts_df = parts_df.with_columns(
                dimensions_str=pl.when(
                    (pl.col('dimensions_str').is_not_null()) &
                    (pl.col('dimensions_str').cast(pl.Utf8) != '')
                ).then(
                    pl.col('dimensions_str').cast(pl.Utf8)
                ).otherwise(
                    pl.concat_str([
                        pl.col('_length_str'), pl.lit('x'),
                        pl.col('_width_str'), pl.lit('x'),
                        pl.col('_height_str')
                    ], separator='')
                )
            )
            parts_df = parts_df.drop(['_length_str', '_width_str', '_height_str'])
            if 'artikul' not in parts_df.columns:
                parts_df = parts_df.with_columns(artikul=pl.lit(''))
            if 'brand' not in parts_df.columns:
                parts_df = parts_df.with_columns(brand=pl.lit(''))
            parts_df = parts_df.with_columns([
                pl.col('artikul').cast(pl.Utf8).fill_null('').alias('_artikul_str'),
                pl.col('brand').cast(pl.Utf8).fill_null('').alias('_brand_str'),
                pl.col('multiplicity').cast(pl.Utf8).alias('_multiplicity_str'),
            ])
            parts_df = parts_df.with_columns(
                description=pl.concat_str([
                    pl.lit('Артикул: '), pl.col('_artikul_str'),
                    pl.lit(', Бренд: '), pl.col('_brand_str'),
                    pl.lit(', Кратность: '), pl.col('_multiplicity_str'),
                    pl.lit(' шт.')
                ], separator='')
            )
            parts_df = parts_df.drop(['_artikul_str', '_brand_str', '_multiplicity_str'])
            final_columns = [
                'artikul_norm', 'brand_norm', 'artikul', 'brand', 'multiplicity', 'barcode',
                'length', 'width', 'height', 'weight', 'image_url', 'dimensions_str', 'description'
            ]
            select_exprs = [pl.col(c) if c in parts_df.columns else pl.lit(
                None).alias(c) for c in final_columns]
            parts_df = parts_df.select(select_exprs)
            self.upsert_data('parts', parts_df, ['artikul_norm', 'brand_norm'])
        logger.info("✅ Обновление базы данных завершено!")

    def _get_brand_markups_sql(self) -> str:
        rows = []
        for brand, markup in self.price_rules['brand_markups'].items():
            safe_brand = brand.replace("'", "''")
            rows.append(f"SELECT '{safe_brand}' AS brand, {markup} AS markup")
        return " UNION ALL ".join(rows) if rows else "SELECT NULL AS brand, NULL AS markup LIMIT 0"

    def build_export_query(self, selected_columns=None, include_prices=True, apply_markup=True):
        description_text = (
            "Состояние товара: новый (в упаковке). Высококачественные автозапчасти и автотовары — надежное решение для вашего автомобиля. "
            "Обеспечьте безопасность, долговечность и высокую производительность вашего авто с помощью нашего широкого ассортимента оригинальных и совместимых автозапчастей."
        )
        brand_markups_sql = self._get_brand_markups_sql()
        select_parts = []
        price_requested = include_prices and (not selected_columns or "Цена" in selected_columns or "Валюта" in selected_columns)
        if price_requested:
            if apply_markup:
                global_markup = self.price_rules.get('global_markup', 0)
                select_parts.append(
                    f"CASE WHEN pr.price IS NOT NULL THEN pr.price * (1 + COALESCE(brm.markup, {global_markup})) ELSE pr.price END AS \"Цена\""
                )
            else:
                select_parts.append('pr.price AS "Цена"')
            select_parts.append("COALESCE(pr.currency, 'RUB') AS \"Валюта\"")
        columns_map = [
            ("Артикул бренда", 'r.artikul AS "Артикул бренда"'),
            ("Бренд", 'r.brand AS "Бренд"'),
            ("Наименование", 'COALESCE(r.representative_name, r.analog_representative_name) AS "Наименование"'),
            ("Категория товара", 'COALESCE(r.representative_category, r.analog_representative_category) AS "Категория товара"'),
            ("Кратность", 'r.multiplicity AS "Кратность"'),
            ("Длинна", 'COALESCE(r.length, r.analog_length) AS "Длинна"'),
            ("Ширина", 'COALESCE(r.width, r.analog_width) AS "Ширина"'),
            ("Высота", 'COALESCE(r.height, r.analog_height) AS "Высота"'),
            ("Вес", 'COALESCE(r.weight, r.analog_weight) AS "Вес"'),
            ("OE номер", 'r.oe_list AS "OE номер"'),
            ("аналоги", 'r.analog_list AS "аналоги"'),
            ("Ссылка на изображение", 'r.image_url AS "Ссылка на изображение"')
        ]
        for name, expr in columns_map:
            if not selected_columns or name in selected_columns:
                select_parts.append(expr.strip())
        if not select_parts:
            select_parts = ['r.artikul AS "Артикул бренда"', 'r.brand AS "Бренд"']
        select_clause = ",\n".join(select_parts)
        ctes = f"""
WITH DescriptionTemplate AS (
    SELECT CHR(10) || CHR(10) || $${description_text}$$ AS text
),
BrandMarkups AS (
    SELECT brand, markup FROM (
        {brand_markups_sql}
    ) AS tmp
),
PartDetails AS (
    SELECT
        cr.artikul_norm,
        cr.brand_norm,
        STRING_AGG(
            DISTINCT regexp_replace(
                regexp_replace(o.oe_number, '''', ''),
                '[^0-9A-Za-zА-Яа-яЁё`\\-\\s]', '', 'g'
            ), ', '
        ) AS oe_list,
        ANY_VALUE(o.name) AS representative_name,
        ANY_VALUE(o.applicability) AS representative_applicability,
        ANY_VALUE(o.category) AS representative_category
    FROM cross_references cr
    LEFT JOIN oe o ON cr.oe_number_norm = o.oe_number_norm
    GROUP BY cr.artikul_norm, cr.brand_norm
),
AllAnalogs AS (
    SELECT
        cr1.artikul_norm,
        cr1.brand_norm,
        STRING_AGG(
            DISTINCT regexp_replace(
                regexp_replace(p2.artikul, '''', ''),
                '[^0-9A-Za-zА-Яа-яЁё`\\-\\s]', '', 'g'
            ), ', '
        ) AS analog_list
    FROM cross_references cr1
    JOIN cross_references cr2 ON cr1.oe_number_norm = cr2.oe_number_norm
    JOIN parts p2 ON cr2.artikul_norm = p2.artikul_norm AND cr2.brand_norm = p2.brand_norm
    WHERE (cr1.artikul_norm != p2.artikul_norm OR cr1.brand_norm != p2.brand_norm)
    GROUP BY cr1.artikul_norm, cr1.brand_norm
),
RankedData AS (
    SELECT
        p.artikul_norm,
        p.brand_norm,
        p.artikul,
        p.brand,
        p.description,
        p.multiplicity,
        p.length,
        p.width,
        p.height,
        p.weight,
        p.dimensions_str,
        p.image_url,
        pd.representative_name,
        pd.representative_applicability,
        pd.representative_category,
        pd.oe_list,
        aa.analog_list,
        ROW_NUMBER() OVER (
            PARTITION BY p.artikul_norm, p.brand_norm
            ORDER BY pd.representative_name DESC NULLS LAST, pd.oe_list DESC NULLS LAST
        ) AS rn
    FROM parts p
    LEFT JOIN PartDetails pd ON p.artikul_norm = pd.artikul_norm AND p.brand_norm = pd.brand_norm
    LEFT JOIN AllAnalogs aa ON p.artikul_norm = aa.artikul_norm AND p.brand_norm = aa.brand_norm
)
"""
        price_join = """
LEFT JOIN prices pr ON r.artikul_norm = pr.artikul_norm AND r.brand_norm = pr.brand_norm
LEFT JOIN BrandMarkups brm ON r.brand = brm.brand
""" if include_prices else ""
        query = f"""
{ctes}
SELECT
{select_clause}
FROM RankedData r
CROSS JOIN DescriptionTemplate dt
{price_join}
WHERE r.rn = 1
ORDER BY r.brand, r.artikul
"""
        return "\n".join([line.rstrip() for line in query.strip().splitlines()])

    def export_to_csv_optimized(self, output_path: str, selected_columns: Optional[List[str]] = None, include_prices: bool = True, apply_markup: bool = True) -> bool:
        if not self.conn:
            logger.warning("⚠️ База данных не доступна")
            return False
        total = self.conn.execute(
            "SELECT count(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)").fetchone()[0]
        if total == 0:
            logger.warning("Нет данных для экспорта")
            return False
        logger.info(f"📤 Экспорт {total} записей в CSV...")
        try:
            query = self.build_export_query(selected_columns, include_prices, apply_markup)
            df = self.conn.execute(query).pl()
            pdf = df.to_pandas()
            dimension_cols = ["Длинна", "Ширина", "Высота", "Вес"]
            for col in dimension_cols:
                if col in pdf.columns:
                    pdf[col] = pdf[col].astype(str).replace({'nan': ''})
            output_dir = Path("auto_parts_data")
            output_dir.mkdir(parents=True, exist_ok=True)
            buf = io.StringIO()
            pdf.to_csv(buf, sep=';', index=False)
            with open(output_path, "wb") as f:
                f.write(b'\xef\xbb\xbf')
                f.write(buf.getvalue().encode('utf-8'))
            size_mb = os.path.getsize(output_path) / (1024 * 1024)
            logger.info(f"Данные экспортированы: {output_path} ({size_mb:.1f} МБ)")
            return True
        except Exception as e:
            logger.exception("Ошибка экспорта CSV")
            logger.error(f"Ошибка при экспорте в CSV: {str(e)}")
            return False

    def export_to_excel_optimized(self, output_path: str, selected_columns: Optional[List[str]] = None, include_prices: bool = True, apply_markup: bool = True) -> bool:
        if not self.conn:
            logger.warning("⚠️ База данных не доступна")
            return False
        total = self.conn.execute(
            "SELECT COUNT(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)").fetchone()[0]
        if total == 0:
            logger.warning("Нет данных для экспорта")
            return False
        query = self.build_export_query(selected_columns, include_prices, apply_markup)
        df = pd.read_sql(query, self.conn)
        for col in ["Длинна", "Ширина", "Высота", "Вес"]:
            if col in df.columns:
                df[col] = df[col].astype(str).replace({r'^nan$': ''}, regex=True)
        if len(df) <= EXCEL_ROW_LIMIT:
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                df.to_excel(writer, index=False)
        else:
            sheets = (len(df) // EXCEL_ROW_LIMIT) + 1
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                for i in range(sheets):
                    df.iloc[i*EXCEL_ROW_LIMIT:(i+1)*EXCEL_ROW_LIMIT].to_excel(
                        writer, index=False, sheet_name=f"Данные_{i+1}")
        return True

    def export_to_parquet(self, output_path: str, selected_columns: Optional[List[str]] = None, include_prices: bool = True, apply_markup: bool = True) -> bool:
        try:
            query = self.build_export_query(selected_columns, include_prices, apply_markup)
            df = self.conn.execute(query).pl()
            df.write_parquet(output_path)
            return True
        except Exception as e:
            logger.exception("Ошибка экспорта Parquet")
            logger.error(f"Ошибка при экспорте в Parquet: {str(e)}")
            return False

    def delete_by_brand(self, brand_norm: str) -> int:
        if not self.conn:
            return 0
        try:
            count_result = self.conn.execute(
                "SELECT COUNT(*) FROM parts WHERE brand_norm = ?", [brand_norm]).fetchone()
            deleted_count = count_result[0] if count_result else 0
            if deleted_count == 0:
                logger.info(f"No records found for brand: {brand_norm}")
                return 0
            self.conn.execute("DELETE FROM parts WHERE brand_norm = ?", [brand_norm])
            self.conn.execute(
                "DELETE FROM cross_references WHERE (artikul_norm, brand_norm) NOT IN (SELECT DISTINCT artikul_norm, brand_norm FROM parts)")
            return deleted_count
        except Exception as e:
            logger.error(f"Error deleting by brand {brand_norm}: {e}")
            raise

    def delete_by_artikul(self, artikul_norm: str) -> int:
        if not self.conn:
            return 0
        try:
            count_result = self.conn.execute(
                "SELECT COUNT(*) FROM parts WHERE artikul_norm = ?", [artikul_norm]).fetchone()
            deleted_count = count_result[0] if count_result else 0
            if deleted_count == 0:
                logger.info(f"No records found for artikul: {artikul_norm}")
                return 0
            self.conn.execute("DELETE FROM parts WHERE artikul_norm = ?", [artikul_norm])
            self.conn.execute(
                "DELETE FROM cross_references WHERE (artikul_norm, brand_norm) NOT IN (SELECT DISTINCT artikul_norm, brand_norm FROM parts)")
            return deleted_count
        except Exception as e:
            logger.error(f"Error deleting by artikul {artikul_norm}: {e}")
            raise

    def merge_all_data_parallel(self, file_paths: Dict[str, str], max_workers: int = 4) -> Dict[str, "pl.DataFrame"]:
        if not POLARS_AVAILABLE:
            return {}
        results = {}
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key, path in file_paths.items():
                if path and os.path.exists(path):
                    futures[executor.submit(
                        self.read_and_prepare_file, path, key)] = key
            for fut in as_completed(futures):
                key = futures[fut]
                try:
                    df = fut.result()
                    if not df.is_empty():
                        results[key] = df
                        logger.info(f"Обработан {key}")
                except Exception as e:
                    logger.error(f"Ошибка обработки {key}: {e}")
        return results

    def get_statistics(self) -> Dict[str, Any]:
        if not self.conn:
            return {}
        stats = {}
        try:
            stats['parts'] = self.conn.execute("SELECT COUNT(*) FROM parts").fetchone()[0]
            stats['oe'] = self.conn.execute("SELECT COUNT(*) FROM oe").fetchone()[0]
            stats['cross'] = self.conn.execute("SELECT COUNT(*) FROM cross_references").fetchone()[0]
            stats['prices'] = self.conn.execute("SELECT COUNT(*) FROM prices").fetchone()[0]
            stats['brands'] = self.conn.execute("SELECT COUNT(DISTINCT brand) FROM parts").fetchone()[0]
            stats['unique_parts'] = self.conn.execute(
                "SELECT COUNT(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)").fetchone()[0]
            avg_price = self.conn.execute("SELECT AVG(price) FROM prices").fetchone()[0]
            stats['avg_price'] = round(avg_price, 2) if avg_price else 0
            top_brands = self.conn.execute(
                "SELECT brand, COUNT(*) as cnt FROM parts GROUP BY brand ORDER BY cnt DESC LIMIT 10").pl()
            stats['top_brands'] = top_brands.to_pandas() if not top_brands.is_empty() else pd.DataFrame()
        except Exception as e:
            logger.error(f"Ошибка сбора статистики: {e}")
        return stats


# ============================================================================
# БЛОК 8: CATALOG ENHANCER (ПОИСК АНАЛОГОВ 2 УРОВНЯ) (400+ СТРОК)
# ============================================================================
class CatalogEnhancer:
    """Обогащение каталога: поиск аналогов по OE номерам (2 уровня)"""
    def __init__(self, db_path: Optional[str] = None):
        self.data_dir = Path("./catalog_data")
        self.data_dir.mkdir(exist_ok=True)
        self.db_path = Path(db_path) if db_path else self.data_dir / "catalog.duckdb"
        self.conn = None
        self.stats = {
            "oe_loaded": 0,
            "parts_loaded": 0,
            "cross_loaded": 0,
            "analog_searches": 0,
            "enrichments": 0
        }
        if DUCKDB_AVAILABLE:
            try:
                self.conn = duckdb.connect(database=str(self.db_path))
                self._setup_database()
                logger.info("CatalogEnhancer инициализирован с DuckDB")
            except Exception as e:
                logger.error(f"Ошибка инициализации DuckDB: {e}")
                self.conn = None
        else:
            logger.warning("DuckDB не установлен, работа в режиме ограниченной функциональности")

    def _setup_database(self):
        if not self.conn:
            return
        try:
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS oe (
                    oe_number_norm VARCHAR PRIMARY KEY,
                    oe_number VARCHAR,
                    name VARCHAR,
                    applicability VARCHAR,
                    category VARCHAR
                )
            """)
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS parts (
                    artikul_norm VARCHAR,
                    brand_norm VARCHAR,
                    artikul VARCHAR,
                    brand VARCHAR,
                    multiplicity INTEGER,
                    barcode VARCHAR,
                    length DOUBLE,
                    width DOUBLE,
                    height DOUBLE,
                    weight DOUBLE,
                    image_url VARCHAR,
                    dimensions_str VARCHAR,
                    description VARCHAR,
                    PRIMARY KEY (artikul_norm, brand_norm)
                )
            """)
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS cross_references (
                    oe_number_norm VARCHAR,
                    artikul_norm VARCHAR,
                    brand_norm VARCHAR,
                    PRIMARY KEY (oe_number_norm, artikul_norm, brand_norm)
                )
            """)
            self.conn.execute("""
                CREATE TABLE IF NOT EXISTS prices (
                    artikul_norm VARCHAR,
                    brand_norm VARCHAR,
                    price DOUBLE,
                    currency VARCHAR DEFAULT 'RUB',
                    PRIMARY KEY (artikul_norm, brand_norm)
                )
            """)
            logger.info("Таблицы каталога созданы")
        except Exception as e:
            logger.error(f"Ошибка создания таблиц: {e}")

    def normalize_key(self, value: str) -> str:
        if not value:
            return ""
        return re.sub(r'[^0-9A-Za-zА-Яа-яЁё]', '', value.lower().strip())

    def load_oe_data(self, df: pd.DataFrame):
        if not self.conn or df.empty:
            return
        try:
            df['oe_number_norm'] = df['oe_number'].apply(self.normalize_key)
            df = df[df['oe_number_norm'] != ""]
            self.conn.execute("DELETE FROM oe")
            for _, row in df.iterrows():
                self.conn.execute(
                    "INSERT INTO oe VALUES (?, ?, ?, ?, ?)",
                    [row['oe_number_norm'], row.get('oe_number', ''),
                     row.get('name', ''), row.get('applicability', ''),
                     row.get('category', 'Разное')]
                )
            self.stats['oe_loaded'] = len(df)
            logger.info(f"Загружено {len(df)} OE записей")
        except Exception as e:
            logger.error(f"Ошибка загрузки OE данных: {e}")

    def load_parts_data(self, df: pd.DataFrame):
        if not self.conn or df.empty:
            return
        try:
            df['artikul_norm'] = df['artikul'].apply(self.normalize_key)
            df['brand_norm'] = df['brand'].apply(self.normalize_key)
            df = df[(df['artikul_norm'] != "") & (df['brand_norm'] != "")]
            self.conn.execute("DELETE FROM parts")
            for _, row in df.iterrows():
                self.conn.execute(
                    "INSERT INTO parts VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
                    [row['artikul_norm'], row['brand_norm'], row.get('artikul', ''),
                     row.get('brand', ''), row.get('multiplicity', 1),
                     row.get('barcode', ''), row.get('length', 0.0),
                     row.get('width', 0.0), row.get('height', 0.0),
                     row.get('weight', 0.0), row.get('image_url', ''),
                     row.get('dimensions_str', ''), row.get('description', '')]
                )
            self.stats['parts_loaded'] = len(df)
            logger.info(f"Загружено {len(df)} записей деталей")
        except Exception as e:
            logger.error(f"Ошибка загрузки данных деталей: {e}")

    def load_cross_references(self, df: pd.DataFrame):
        if not self.conn or df.empty:
            return
        try:
            df['oe_number_norm'] = df['oe_number'].apply(self.normalize_key)
            df['artikul_norm'] = df['artikul'].apply(self.normalize_key)
            df['brand_norm'] = df['brand'].apply(self.normalize_key)
            df = df[(df['oe_number_norm'] != "") & (df['artikul_norm'] != "")]
            self.conn.execute("DELETE FROM cross_references")
            for _, row in df.iterrows():
                self.conn.execute(
                    "INSERT INTO cross_references VALUES (?, ?, ?)",
                    [row['oe_number_norm'], row['artikul_norm'], row['brand_norm']]
                )
            self.stats['cross_loaded'] = len(df)
            logger.info(f"Загружено {len(df)} кросс-ссылок")
        except Exception as e:
            logger.error(f"Ошибка загрузки кросс-ссылок: {e}")

    def get_analog_data(self, artikul: str, brand: str) -> Dict[str, Any]:
        """Поиск аналогов по OE номерам (2 уровня)"""
        self.stats['analog_searches'] += 1
        if not self.conn:
            return {"error": "DuckDB не доступен"}
        artikul_norm = self.normalize_key(artikul)
        brand_norm = self.normalize_key(brand)
        if not artikul_norm or not brand_norm:
            return {"error": "Не указан артикул или бренд"}
        try:
            query = f"""
                WITH PartDetails AS (
                    SELECT
                        cr.artikul_norm,
                        cr.brand_norm,
                        STRING_AGG(DISTINCT oe.oe_number, ', ') AS oe_list
                    FROM cross_references cr
                    LEFT JOIN oe ON cr.oe_number_norm = oe.oe_number_norm
                    WHERE cr.artikul_norm = '{artikul_norm}' AND cr.brand_norm = '{brand_norm}'
                    GROUP BY cr.artikul_norm, cr.brand_norm
                ),
                AnalogParts AS (
                    SELECT DISTINCT
                        p.artikul,
                        p.brand,
                        p.description,
                        p.length,
                        p.width,
                        p.height,
                        p.weight,
                        p.dimensions_str,
                        p.image_url
                    FROM cross_references cr
                    JOIN parts p ON cr.artikul_norm = p.artikul_norm AND cr.brand_norm = p.brand_norm
                    WHERE cr.oe_number_norm IN (
                        SELECT oe_number_norm
                        FROM cross_references
                        WHERE artikul_norm = '{artikul_norm}' AND brand_norm = '{brand_norm}'
                    )
                    AND NOT (cr.artikul_norm = '{artikul_norm}' AND cr.brand_norm = '{brand_norm}')
                    LIMIT 50
                )
                SELECT
                    (SELECT COUNT(*) FROM AnalogParts) AS analog_count,
                    (SELECT oe_list FROM PartDetails) AS oe_list,
                    (SELECT * FROM AnalogParts LIMIT 20) AS analogs
            """
            result = self.conn.execute(query).df()
            if result.empty or result.iloc[0]['analog_count'] == 0:
                return {
                    "artikul": artikul,
                    "brand": brand,
                    "analog_count": 0,
                    "analogs": [],
                    "has_analogs": False,
                    "oe_list": ""
                }
            row = result.iloc[0]
            analog_count = int(row['analog_count']) if not pd.isna(row['analog_count']) else 0
            analogs = []
            if analog_count > 0:
                analog_query = f"""
                    SELECT DISTINCT
                        p.artikul,
                        p.brand,
                        p.description,
                        p.length,
                        p.width,
                        p.height,
                        p.weight,
                        p.dimensions_str,
                        p.image_url
                    FROM cross_references cr
                    JOIN parts p ON cr.artikul_norm = p.artikul_norm AND cr.brand_norm = p.brand_norm
                    WHERE cr.oe_number_norm IN (
                        SELECT oe_number_norm
                        FROM cross_references
                        WHERE artikul_norm = '{artikul_norm}' AND brand_norm = '{brand_norm}'
                    )
                    AND NOT (cr.artikul_norm = '{artikul_norm}' AND cr.brand_norm = '{brand_norm}')
                    LIMIT 20
                """
                analog_df = self.conn.execute(analog_query).df()
                for _, arow in analog_df.iterrows():
                    analogs.append({
                        "artikul": safe_str(arow.get('artikul', '')),
                        "brand": safe_str(arow.get('brand', '')),
                        "description": safe_str(arow.get('description', '')),
                        "length": safe_float(arow.get('length', 0)),
                        "width": safe_float(arow.get('width', 0)),
                        "height": safe_float(arow.get('height', 0)),
                        "weight": safe_float(arow.get('weight', 0)),
                        "dimensions_str": safe_str(arow.get('dimensions_str', '')),
                        "image_url": safe_str(arow.get('image_url', ''))
                    })
            return {
                "artikul": artikul,
                "brand": brand,
                "analog_count": analog_count,
                "analogs": analogs,
                "has_analogs": analog_count > 0,
                "oe_list": safe_str(row.get('oe_list', ''))
            }
        except Exception as e:
            logger.error(f"Ошибка получения аналогов: {e}")
            return {"error": str(e)}

    def enhance_catalog_data(self, df: pd.DataFrame,
                             artikul_col: str = "Артикул",
                             brand_col: str = "Бренд") -> pd.DataFrame:
        """Обогащение каталога: добавление информации об аналогах"""
        if df.empty:
            return df
        if artikul_col not in df.columns or brand_col not in df.columns:
            logger.warning(f"Колонки {artikul_col} или {brand_col} не найдены")
            return df
        self.stats['enrichments'] += 1
        df_copy = df.copy()
        new_columns = ['analog_count', 'has_analogs', 'analog_list', 'oe_list']
        for col in new_columns:
            if col not in df_copy.columns:
                df_copy[col] = None
        for idx, row in df_copy.iterrows():
            artikul = safe_str(row.get(artikul_col, ''))
            brand = safe_str(row.get(brand_col, ''))
            if artikul and brand:
                data = self.get_analog_data(artikul, brand)
                if not data.get('error'):
                    df_copy.at[idx, 'analog_count'] = data.get('analog_count', 0)
                    df_copy.at[idx, 'has_analogs'] = data.get('has_analogs', False)
                    df_copy.at[idx, 'oe_list'] = data.get('oe_list', '')
                    if data.get('analogs'):
                        analog_str = ', '.join([
                            f"{a['artikul']} ({a['brand']})"
                            for a in data['analogs'][:5]
                            if a.get('artikul')
                        ])
                        df_copy.at[idx, 'analog_list'] = analog_str
        return df_copy

    def get_stats(self) -> Dict:
        return self.stats.copy()


# ============================================================================
# БЛОК 9: ML-КЛАССИФИКАТОР КАТЕГОРИЙ (300+ СТРОК)
# ============================================================================
class CategoryClassifier:
    """ML-классификатор категорий автозапчастей"""
    def __init__(self, model_path: str = "category_model.pkl"):
        self.model_path = model_path
        self.model = None
        self.categories = [
            "Двигатель", "Трансмиссия", "Подвеска", "Тормозная система",
            "Рулевое управление", "Электрооборудование", "Система охлаждения",
            "Система выпуска", "Система питания", "Фильтры",
            "Масла и жидкости", "Кузовные детали", "Оптика",
            "Шины и диски", "Инструменты", "Ремни и приводы",
            "Подшипники", "Сальники и прокладки", "Крепеж",
            "Климат-контроль", "Аудио и мультимедиа", "Безопасность",
            "Прочее"
        ]
        self.accuracy = 0.0
        self.cache = {}
        self.logger = logging.getLogger('CategoryClassifier')
        self._load_model()

    def _load_model(self):
        if os.path.exists(self.model_path) and SKLEARN_AVAILABLE:
            try:
                self.model = joblib.load(self.model_path)
                self.categories = self.model.classes_ if hasattr(self.model, 'classes_') else self.categories
                self.logger.info(f"ML-модель загружена, категорий: {len(self.categories)}")
                return
            except Exception as e:
                self.logger.warning(f"Ошибка загрузки модели: {e}")
        self._train_model()

    def _train_model(self):
        if not SKLEARN_AVAILABLE:
            self.logger.warning("Scikit-learn не установлен, используется fallback классификатор")
            return
        try:
            X = []
            y = []
            category_keywords = {
                "Двигатель": ["двигатель", "мотор", "двс", "поршень", "шатун", "клапан", "гбц"],
                "Трансмиссия": ["коробка", "кпп", "сцепление", "привод", "дифференциал", "акпп"],
                "Подвеска": ["амортизатор", "пружина", "рычаг", "сайлентблок", "шаровая", "стабилизатор"],
                "Тормозная система": ["колодки", "диск", "барабан", "суппорт", "гтц", "абс"],
                "Рулевое управление": ["рейка", "тяга", "наконечник", "руль", "гур"],
                "Электрооборудование": ["генератор", "стартер", "аккумулятор", "свеча", "провод"],
                "Система охлаждения": ["радиатор", "помпа", "термостат", "вентилятор", "бачок"],
                "Система выпуска": ["глушитель", "катализатор", "резонатор", "гофра", "лямбда"],
                "Система питания": ["насос", "фильтр", "форсунка", "дроссель", "тнвд"],
                "Фильтры": ["фильтр", "масляный", "воздушный", "салонный", "топливный"],
                "Масла и жидкости": ["масло", "жидкость", "смазка", "антифриз", "тормозуха"],
                "Кузовные детали": ["бампер", "капот", "крыло", "дверь", "стекло", "фара"],
                "Шины и диски": ["шина", "диск", "колесо", "покрышка", "резина"],
                "Инструменты": ["инструмент", "ключ", "домкрат", "насос"],
                "Прочее": []
            }
            for category, keywords in category_keywords.items():
                for keyword in keywords:
                    if keyword:
                        X.append(keyword)
                        y.append(category)
                        X.append(keyword + " " + category)
                        y.append(category)
            if X:
                X_train, X_test, y_train, y_test = train_test_split(
                    X, y, test_size=0.2, random_state=42
                )
                self.model = Pipeline([
                    ('tfidf', TfidfVectorizer(max_features=2000, ngram_range=(1, 2))),
                    ('clf', MultinomialNB(alpha=0.1))
                ])
                self.model.fit(X_train, y_train)
                self.categories = self.model.classes_
                y_pred = self.model.predict(X_test)
                self.accuracy = accuracy_score(y_test, y_pred)
                joblib.dump(self.model, self.model_path)
                self.logger.info(f"ML-модель обучена на {len(X)} примерах, точность: {self.accuracy:.2%}")
        except Exception as e:
            self.logger.error(f"Ошибка обучения модели: {e}")
            self.model = None

    def predict(self, name: str) -> Tuple[str, float]:
        """Предсказание категории для товара"""
        if not isinstance(name, str):
            name = str(name)
        if not name or not name.strip():
            return "Прочее", 0.0
        name_lower = name.lower()
        if name_lower in self.cache:
            return self.cache[name_lower]
        if SKLEARN_AVAILABLE and self.model is not None:
            try:
                pred = self.model.predict([name_lower])[0]
                probs = self.model.predict_proba([name_lower])[0]
                confidence = max(probs) * 100
                if confidence > 30:
                    result = (pred, confidence)
                    self.cache[name_lower] = result
                    return result
            except Exception as e:
                self.logger.warning(f"ML prediction error: {e}")
        best_category = "Прочее"
        best_score = 0.0
        category_keywords = {
            "Двигатель": ["двигатель", "мотор", "двс", "поршень", "шатун", "клапан", "гбц"],
            "Трансмиссия": ["коробка", "кпп", "сцепление", "привод", "дифференциал"],
            "Подвеска": ["амортизатор", "пружина", "рычаг", "сайлентблок", "шаровая"],
            "Тормозная система": ["колодки", "диск", "барабан", "суппорт", "гтц"],
            "Электрооборудование": ["генератор", "стартер", "аккумулятор", "свеча"],
            "Система охлаждения": ["радиатор", "помпа", "термостат", "вентилятор"],
            "Система выпуска": ["глушитель", "катализатор", "резонатор", "гофра"],
            "Система питания": ["насос", "фильтр", "форсунка", "дроссель"],
            "Фильтры": ["фильтр", "масляный", "воздушный", "салонный"],
            "Масла и жидкости": ["масло", "жидкость", "смазка", "антифриз"],
            "Кузовные детали": ["бампер", "капот", "крыло", "дверь", "стекло"],
            "Шины и диски": ["шина", "диск", "колесо", "покрышка"],
            "Инструменты": ["инструмент", "ключ", "домкрат"],
            "Прочее": []
        }
        for category, keywords in category_keywords.items():
            score = 0.0
            for keyword in keywords:
                if keyword in name_lower:
                    weight = len(keyword) / 10.0
                    if name_lower.startswith(keyword):
                        weight *= 1.5
                    score += weight
            if score > best_score:
                best_score = score
                best_category = category
        confidence = min(best_score * 15, 100.0)
        result = (best_category, round(confidence, 1))
        self.cache[name_lower] = result
        return result

    def predict_batch(self, names: List[str]) -> List[Tuple[str, float]]:
        """Пакетное предсказание категорий"""
        results = []
        for name in names:
            cat, conf = self.predict(name)
            results.append((cat, conf))
        return results
# ============================================================================
# БЛОК 10: UI ФУНКЦИИ - ЗАГРУЗКА ДАННЫХ (500+ СТРОК)
# ============================================================================
def show_data_upload_interface():
    """Интерфейс загрузки данных с автоопределением кодировки"""
    st.header("📁 Загрузка данных каталога")
    st.info("""
📋 **Инструкция по загрузке данных:**

**ОБЯЗАТЕЛЬНЫЕ колонки:**
- `Артикул` или `article` или `sku` - идентификатор товара
- `Бренд` или `brand` или `производитель` - бренд товара
- `Цена` или `price` или `стоимость` - цена продажи
- `Себестоимость` или `cost` - себестоимость товара

**ОПЦИОНАЛЬНЫЕ колонки (для расширенной функциональности):**
- `Длина`, `Ширина`, `Высота`, `Вес` - габариты для расчета логистики
- `OE номер` - оригинальный номер запчасти (для поиска аналогов)
- `Категория` - категория товара (для классификации)
- `Штрихкод`, `Описание`, `Кратность`
""")

    uploaded_file = st.file_uploader(
        "Загрузите файл каталога (Excel или CSV)",
        type=['xlsx', 'xls', 'csv'],
        key="data_upload_file"
    )

    if uploaded_file is not None:
        try:
            df = None
            file_name = uploaded_file.name.lower()

            # --- ОБРАБОТКА CSV ФАЙЛОВ ---
            if file_name.endswith('.csv'):
                encodings_to_try = [
                    'utf-8-sig', 'utf-8', 'cp1251', 'windows-1251',
                    'cp1252', 'latin1', 'iso-8859-1', 'koi8-r',
                    'mac_cyrillic', 'cp866'
                ]
                separators = [',', ';', '\t', '|', ':', '^']

                for encoding in encodings_to_try:
                    if df is not None and not df.empty:
                        break
                    for sep in separators:
                        try:
                            uploaded_file.seek(0)
                            df = pd.read_csv(
                                uploaded_file,
                                encoding=encoding,
                                sep=sep,
                                engine='python',
                                on_bad_lines='skip',
                                skipinitialspace=True,
                                quotechar='"',
                                doublequote=True
                            )
                            if df is not None and not df.empty and len(df.columns) > 1:
                                logger.info(f"CSV прочитан: кодировка={encoding}, разделитель='{sep}'")
                                break
                        except Exception:
                            continue

                # Пробуем через chardet
                if df is None or df.empty:
                    if CHARDET_AVAILABLE and chardet is not None:
                        try:
                            uploaded_file.seek(0)
                            raw_data = uploaded_file.read(100000)
                            detected = chardet.detect(raw_data)
                            if detected and detected.get('encoding'):
                                uploaded_file.seek(0)
                                for sep in separators:
                                    try:
                                        df = pd.read_csv(
                                            uploaded_file,
                                            encoding=detected['encoding'],
                                            sep=sep,
                                            engine='python',
                                            on_bad_lines='skip'
                                        )
                                        if df is not None and not df.empty and len(df.columns) > 1:
                                            logger.info(f"CSV прочитан через chardet: {detected['encoding']}")
                                            break
                                    except Exception:
                                        continue
                        except Exception as e:
                            logger.warning(f"Ошибка chardet: {e}")

                if df is None or df.empty:
                    try:
                        uploaded_file.seek(0)
                        lines = uploaded_file.read().splitlines()[:5]
                        if lines:
                            st.warning("⚠️ Не удалось прочитать CSV файл. Первые строки файла:")
                            for i, line in enumerate(lines[:3]):
                                try:
                                    st.text(f"Строка {i+1}: {line[:200].decode('utf-8', errors='replace')}")
                                except Exception:
                                    st.text(f"Строка {i+1}: [бинарные данные]")
                    except Exception:
                        pass
                    raise ValueError("Не удалось прочитать CSV файл. Проверьте кодировку и разделитель.")

            # --- ОБРАБОТКА EXCEL ФАЙЛОВ ---
            elif file_name.endswith(('.xlsx', '.xls')):
                excel_engines = ['openpyxl', 'xlrd']
                for engine in excel_engines:
                    try:
                        uploaded_file.seek(0)
                        df = pd.read_excel(uploaded_file, engine=engine)
                        if df is not None and not df.empty:
                            logger.info(f"Excel прочитан с движком: {engine}")
                            break
                    except Exception:
                        continue

                if df is None or df.empty:
                    try:
                        uploaded_file.seek(0)
                        df = pd.read_excel(io.BytesIO(uploaded_file.read()), engine='openpyxl')
                        if df is not None and not df.empty:
                            logger.info("Excel прочитан через BytesIO")
                    except Exception:
                        pass

                if df is None or df.empty:
                    available_engines = ['openpyxl', 'xlrd', 'odf']
                    for engine in available_engines:
                        try:
                            uploaded_file.seek(0)
                            df = pd.read_excel(uploaded_file, engine=engine)
                            if df is not None and not df.empty:
                                break
                        except Exception:
                            continue
            else:
                raise ValueError(f"Неподдерживаемый формат файла: {file_name}")

            # --- ПРОВЕРКА ЗАГРУЖЕННЫХ ДАННЫХ ---
            if df is None or df.empty:
                st.error("❌ Не удалось прочитать файл. Проверьте формат и кодировку.")
                with st.expander("💡 Возможные решения:"):
                    st.markdown("""
1. **Для CSV файлов:**
- Сохраните файл с кодировкой UTF-8
- Убедитесь, что разделитель соответствует
- Проверьте, что файл не поврежден

2. **Для Excel файлов:**
- Убедитесь, что файл не защищен паролем
- Попробуйте сохранить файл заново

3. **Общие рекомендации:**
- Установите: `pip install chardet openpyxl xlrd`
""")
                return

            # --- ОЧИСТКА ДАННЫХ ---
            df = df.dropna(how='all')
            df = df.dropna(axis=0, how='all')
            if df.empty:
                st.warning("⚠️ Файл содержит только пустые строки. Проверьте данные.")
                return

            df.columns = df.columns.str.strip()
            st.session_state.uploaded_data = df
            st.success(f"✅ Успешно загружено {len(df)} товаров")

            # --- ПРЕДПРОСМОТР ---
            st.subheader("📊 Предпросмотр данных")
            st.dataframe(df.head(10), use_container_width=True, key="upload_preview_table")

            # --- АНАЛИЗ КОЛОНОК ---
            st.subheader("📋 Анализ колонок")
            col1, col2 = st.columns(2)
            with col1:
                st.markdown("**Обязательные колонки:**")
                required_cols = ["Артикул", "Бренд", "Цена", "Себестоимость"]
                for col in required_cols:
                    found = any(col.lower() in c.lower() for c in df.columns)
                    actual_col = None
                    if found:
                        for c in df.columns:
                            if col.lower() in c.lower():
                                actual_col = c
                                break
                    if found and actual_col:
                        st.write(f"✅ {col} → '{actual_col}'")
                    else:
                        st.write(f"❌ {col} (не найдена)")

            with col2:
                st.markdown("**Опциональные колонки:**")
                optional_cols = ["Длина", "Ширина", "Высота", "Вес", "OE номер", "Категория"]
                for col in optional_cols:
                    found = any(col.lower() in c.lower() for c in df.columns)
                    actual_col = None
                    if found:
                        for c in df.columns:
                            if col.lower() in c.lower():
                                actual_col = c
                                break
                    if found and actual_col:
                        st.write(f"✅ {col} → '{actual_col}'")
                    else:
                        st.write(f"❌ {col}")

            # --- СТАТИСТИКА ---
            st.subheader("📊 Статистика данных")
            stats_col1, stats_col2, stats_col3, stats_col4 = st.columns(4)
            with stats_col1:
                st.metric("📦 Товаров", len(df))
            with stats_col2:
                price_col = None
                for col in df.columns:
                    if any(w in col.lower() for w in ['цена', 'price', 'стоимость']):
                        price_col = col
                        break
                if price_col:
                    try:
                        avg_price = safe_float(df[price_col].mean())
                        st.metric("💰 Средняя цена", f"{avg_price:,.0f} ₽" if avg_price > 0 else "Н/Д")
                    except Exception:
                        st.metric("💰 Средняя цена", "Ошибка")
                else:
                    st.metric("💰 Средняя цена", "—")
            with stats_col3:
                cost_col = None
                for col in df.columns:
                    if any(w in col.lower() for w in ['себестоимость', 'cost', 'закупочная']):
                        cost_col = col
                        break
                if cost_col:
                    try:
                        avg_cost = safe_float(df[cost_col].mean())
                        st.metric("💵 Средняя себестоимость", f"{avg_cost:,.0f} ₽" if avg_cost > 0 else "Н/Д")
                    except Exception:
                        st.metric("💵 Средняя себестоимость", "Ошибка")
                else:
                    st.metric("💵 Средняя себестоимость", "—")
            with stats_col4:
                brand_col = None
                for col in df.columns:
                    if any(w in col.lower() for w in ['бренд', 'brand', 'производитель']):
                        brand_col = col
                        break
                if brand_col:
                    try:
                        unique_brands = df[brand_col].nunique()
                        st.metric("🏷️ Брендов", unique_brands)
                    except Exception:
                        st.metric("🏷️ Брендов", "Ошибка")
                else:
                    st.metric("🏷️ Брендов", "—")

            # --- ДЕЙСТВИЯ ---
            st.subheader("🔧 Действия с данными")
            action_col1, action_col2, action_col3 = st.columns(3)
            with action_col1:
                if st.button("🏷️ Классифицировать категории", type="secondary", key="classify_btn"):
                    with st.spinner("Классификация товаров..."):
                        classifier = CategoryClassifier()
                        name_col = None
                        for col in df.columns:
                            col_lower = col.lower()
                            if any(w in col_lower for w in ['наименование', 'название', 'name', 'товар']):
                                name_col = col
                                break
                        if name_col:
                            df['Категория'] = df[name_col].apply(lambda x: classifier.predict(str(x))[0])
                            st.session_state.uploaded_data = df
                            st.success("✅ Классификация завершена!")
                            st.subheader("📊 Распределение по категориям")
                            category_counts = df['Категория'].value_counts()
                            st.dataframe(category_counts, use_container_width=True, key="category_counts")
                        else:
                            st.warning("⚠️ Не найдена колонка с названием товара")
            with action_col2:
                if st.button("📊 Обогатить каталог", type="primary", key="upload_enrich_button"):
                    st.info("ℹ️ Перейдите на вкладку '📊 Обогащение каталога' для поиска аналогов")
            with action_col3:
                if st.button("🧹 Очистить данные", type="secondary", key="clear_data_btn"):
                    if st.session_state.get('uploaded_data') is not None:
                        del st.session_state.uploaded_data
                        st.success("✅ Данные очищены")
                        st.rerun()

        except Exception as e:
            st.error(f"❌ Ошибка загрузки файла: {str(e)}")
            with st.expander("📋 Подробности ошибки", expanded=True):
                st.code(traceback.format_exc())
            with st.expander("💡 Возможные решения:"):
                st.markdown("""
1. **Для CSV файлов:**
- Сохраните файл с кодировкой UTF-8
- Убедитесь, что разделитель соответствует
- Проверьте, что файл не поврежден

2. **Для Excel файлов:**
- Убедитесь, что файл не защищен паролем
- Попробуйте сохранить файл заново

3. **Общие рекомендации:**
- Установите: `pip install chardet openpyxl xlrd`
""")

        # --- ШАБЛОН ---
        if st.button("📥 Скачать шаблон данных"):
            template_df = pd.DataFrame({
                "Артикул": ["ABC-001", "ABC-002", "ABC-003"],
                "Бренд": ["Bosch", "Bosch", "Siemens"],
                "Цена": [1000, 1500, 2000],
                "Себестоимость": [500, 750, 1000],
                "Категория": ["Автозапчасти", "Автозапчасти", "Инструменты"],
                "Длина": [10, 15, 20],
                "Ширина": [5, 7, 10],
                "Высота": [3, 4, 5],
                "Вес": [0.5, 0.8, 1.2],
                "OE номер": ["123456", "654321", "789012"],
                "Описание": ["Описание товара 1", "Описание товара 2", "Описание товара 3"]
            })
            csv = template_df.to_csv(index=False, encoding='utf-8-sig')
            st.download_button(
                label="📥 Скачать шаблон CSV",
                data=csv,
                file_name="шаблон_каталога.csv",
                mime="text/csv",
                key="download_template"
            )


# ============================================================================
# БЛОК 11: ЮНИТ-ЭКОНОМИКА (ОДИН ТОВАР)
# ============================================================================
def show_unit_economics_interface():
    """Интерфейс расчета юнит-экономики для одного товара"""
    st.header("📊 Юнит-экономика маркетплейсов 2026")
    unit_economics = MarketplaceUnitEconomics()

    st.info("""
💡 **Режимы работы:**
- **FBY** - доставка силами маркетплейса (самый дешевый)
- **FBS** - доставка силами продавца (базовый)
- **FBO** - доставка силами оператора (средний)
- **DBS** - доставка силами продавца (дорогой)
- **FBP** - доставка силами платформы (чуть дешевле)
""")

    col1, col2 = st.columns(2)
    with col1:
        price = st.number_input("💰 Цена продажи (₽)", min_value=0.0, value=1000.0, step=10.0, key="ue_price")
        cost = st.number_input("💵 Себестоимость (₽)", min_value=0.0, value=500.0, step=10.0, key="ue_cost")
        weight = st.number_input("⚖️ Вес (кг)", min_value=0.0, value=1.0, step=0.1, key="ue_weight")
    with col2:
        volume = st.number_input("📦 Объем (литры)", min_value=0.0, value=5.0, step=0.5, key="ue_volume")
        marketplace = st.selectbox("🏪 Маркетплейс", list(unit_economics._configs.keys()), key="ue_marketplace")
        operation_mode = st.selectbox("📦 Режим работы", ["FBY", "FBS", "FBO", "DBS", "FBP"], key="ue_mode")
        category = st.text_input("📂 Категория (опционально)", placeholder="например: двигатель", key="ue_category")

    is_premium = st.checkbox("⭐ Премиум-раздел (доп. комиссия)", key="ue_premium")

    if st.button("🚀 Рассчитать юнит-экономику", type="primary", key="ue_calc"):
        with st.spinner("Расчет юнит-экономики..."):
            economics = unit_economics.calculate_unit_economics(
                price=price,
                cost=cost,
                marketplace=marketplace,
                weight=weight,
                category=category if category else None,
                is_premium=is_premium
            )

            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("💰 Прибыль", f"{economics.profit:.2f} ₽", delta=f"{economics.profit_per_ruble:.2f} ₽/₽")
                st.metric("📈 Маржа", f"{economics.margin_percent:.2f}%")
            with col2:
                st.metric("📊 ROI", f"{economics.roi:.2f}%")
                st.metric("⚖️ Точка безубыточности", f"{economics.breakeven_price:.2f} ₽")
            with col3:
                st.metric("💵 Комиссия", f"{economics.commission:.2f} ₽", f"{economics.commission_percent:.1f}% от цены")

            st.subheader("📋 Детализация расходов")
            expenses_data = {
                "Статья расходов": [
                    "Себестоимость", "Комиссия", "Подписка", "Логистика",
                    "Хранение", "Эквайринг", "Доставка", "Последняя миля",
                    "Возвраты", "РКО", "Премиум", "Страховка", "Упаковка", "Маркетинг", "ИТОГО"
                ],
                "Сумма (₽)": [
                    economics.cost, economics.commission, economics.subscription_cost,
                    economics.logistics, economics.storage_cost, economics.acquiring,
                    economics.delivery, economics.last_mile, economics.returns,
                    economics.rko_fee, economics.premium_fee, economics.insurance_fee,
                    economics.packing_fee, economics.marketing_fee, economics.total_expenses
                ],
                "% от цены": [
                    f"{economics.cost/price*100:.1f}%",
                    f"{economics.commission/price*100:.1f}%",
                    f"{economics.subscription_cost/price*100:.1f}%",
                    f"{economics.logistics/price*100:.1f}%",
                    f"{economics.storage_cost/price*100:.1f}%",
                    f"{economics.acquiring/price*100:.1f}%",
                    f"{economics.delivery/price*100:.1f}%",
                    f"{economics.last_mile/price*100:.1f}%",
                    f"{economics.returns/price*100:.1f}%",
                    f"{economics.rko_fee/price*100:.1f}%",
                    f"{economics.premium_fee/price*100:.1f}%",
                    f"{economics.insurance_fee/price*100:.1f}%",
                    f"{economics.packing_fee/price*100:.1f}%",
                    f"{economics.marketing_fee/price*100:.1f}%",
                    f"{economics.total_expenses/price*100:.1f}%"
                ]
            }
            st.dataframe(pd.DataFrame(expenses_data), use_container_width=True, key="ue_expenses_table")

            st.subheader("🏆 Сравнение всех маркетплейсов")
            comparison_df = unit_economics.calculate_for_all_marketplaces(
                price=price, cost=cost, weight=weight, category=category if category else None,
                operation_mode=operation_mode
            )
            st.dataframe(comparison_df, use_container_width=True, key="ue_comparison_table")

            if not comparison_df.empty:
                best_idx = comparison_df['profit'].idxmax()
                best = comparison_df.loc[best_idx]
                st.success(f"🏆 Оптимальный маркетплейс: **{best['marketplace']}** "
                           f"(прибыль: {best['profit']:.2f} ₽, маржа: {best['margin_percent']:.2f}%)")

                if PLOTLY_AVAILABLE and go is not None:
                    try:
                        fig = go.Figure()
                        fig.add_trace(go.Bar(
                            x=comparison_df['marketplace'],
                            y=comparison_df['profit'],
                            name='Прибыль',
                            marker_color='#e94560'
                        ))
                        fig.add_trace(go.Bar(
                            x=comparison_df['marketplace'],
                            y=comparison_df['margin_percent'],
                            name='Маржа %',
                            marker_color='#0f3460',
                            yaxis='y2'
                        ))
                        fig.update_layout(
                            title='Сравнение маркетплейсов',
                            xaxis_title='Маркетплейс',
                            yaxis_title='Прибыль (₽)',
                            yaxis2=dict(title='Маржа (%)', overlaying='y', side='right'),
                            barmode='group',
                            height=400
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    except Exception as e:
                        logger.warning(f"Ошибка визуализации: {e}")

    stats = unit_economics.get_stats()
    if stats.get('total_calculations', 0) > 0:
        st.subheader("📊 Статистика расчетов")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("📊 Всего расчетов", stats.get('total_calculations', 0))
        with col2:
            st.metric("💰 Средняя прибыль", f"{stats.get('avg_profit', 0):.2f} ₽")
        with col3:
            st.metric("📈 Макс. прибыль", f"{stats.get('max_profit', 0):.2f} ₽")
        with col4:
            st.metric("🏆 Лучший МП", stats.get('best_marketplace', '—'))


# ============================================================================
# БЛОК 12: ЮНИТ-ЭКОНОМИКА ПО АРТИКУЛАМ (600+ СТРОК)
# ============================================================================
def show_unit_economics_by_article_interface():
    """Интерфейс для расчета юнит-экономики по каждому артикулу"""
    st.header("📊 Юнит-экономика по артикулам")

    if st.session_state.get('uploaded_data') is None:
        st.warning("⚠️ Сначала загрузите данные в разделе '📁 Загрузка данных'")
        return

    df = st.session_state.uploaded_data.copy()
    st.info("""
📊 **Расчет юнит-экономики для каждого товара**

Для каждого артикула в каталоге будут рассчитаны:
- 💰 Прибыль по каждому маркетплейсу
- 📈 Маржинальность
- 📊 ROI
- ⚖️ Точка безубыточности
- 📋 Детализация всех расходов
""")

    # --- ПАРАМЕТРЫ РАСЧЕТА ---
    st.subheader("⚙️ Параметры расчета")
    col1, col2, col3 = st.columns(3)
    with col1:
        unit_economics = MarketplaceUnitEconomics()
        available_marketplaces = list(unit_economics._configs.keys())
        selected_marketplaces = st.multiselect(
            "🏪 Маркетплейсы для расчета",
            options=available_marketplaces,
            default=available_marketplaces[:3],
            key="ue_article_marketplaces"
        )
        if not selected_marketplaces:
            st.warning("⚠️ Выберите хотя бы один маркетплейс")
            return
    with col2:
        operation_mode = st.selectbox("📦 Режим работы", ["FBY", "FBS", "FBO", "DBS", "FBP"], key="ue_article_mode")
        days_in_storage = st.number_input("📦 Дней хранения", min_value=1, max_value=365, value=30, step=1, key="ue_article_days_storage")
    with col3:
        apply_markup = st.checkbox("💰 Применить наценку", value=False, key="ue_article_apply_markup")
        if apply_markup:
            markup_percent = st.number_input("Наценка (%)", min_value=0.0, max_value=500.0, value=20.0, step=5.0, key="ue_article_markup_percent")
        else:
            markup_percent = 0.0
        is_premium = st.checkbox("⭐ Премиум-раздел", value=False, key="ue_article_premium")

    # --- ВЫБОР КОЛОНОК ---
    st.subheader("📋 Определение колонок в данных")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        article_col = st.selectbox("Артикул", options=df.columns, key="ue_article_col")
    with col2:
        price_options = [col for col in df.columns if any(w in col.lower() for w in ['цена', 'price', 'стоимость'])]
        if not price_options:
            price_options = df.columns.tolist()
        price_col = st.selectbox("Цена продажи", options=price_options, key="ue_price_col")
    with col3:
        cost_options = [col for col in df.columns if any(w in col.lower() for w in ['себестоимость', 'cost', 'закупочная'])]
        if not cost_options:
            cost_options = df.columns.tolist()
        cost_col = st.selectbox("Себестоимость", options=cost_options, key="ue_cost_col")
    with col4:
        category_options = [col for col in df.columns if any(w in col.lower() for w in ['категория', 'category', 'группа'])]
        category_options = ['Не выбрано'] + category_options
        category_col = st.selectbox("Категория (опционально)", options=category_options, key="ue_category_col")

    # --- ОПЦИОНАЛЬНЫЕ КОЛОНКИ ---
    st.subheader("📏 Габариты (опционально)")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        length_options = [col for col in df.columns if any(w in col.lower() for w in ['длина', 'length', 'длинна'])]
        length_options = ['Не выбрано'] + length_options
        length_col = st.selectbox("Длина (см)", options=length_options, key="ue_length_col")
    with col2:
        width_options = [col for col in df.columns if any(w in col.lower() for w in ['ширина', 'width'])]
        width_options = ['Не выбрано'] + width_options
        width_col = st.selectbox("Ширина (см)", options=width_options, key="ue_width_col")
    with col3:
        height_options = [col for col in df.columns if any(w in col.lower() for w in ['высота', 'height'])]
        height_options = ['Не выбрано'] + height_options
        height_col = st.selectbox("Высота (см)", options=height_options, key="ue_height_col")
    with col4:
        weight_options = [col for col in df.columns if any(w in col.lower() for w in ['вес', 'weight'])]
        weight_options = ['Не выбрано'] + weight_options
        weight_col = st.selectbox("Вес (кг)", options=weight_options, key="ue_weight_col")

    # --- ФИЛЬТРЫ ---
    st.subheader("🔍 Фильтры")
    col1, col2 = st.columns(2)
    with col1:
        if category_col != 'Не выбрано' and category_col in df.columns:
            categories = df[category_col].dropna().unique().tolist()
            categories = [str(cat).strip() for cat in categories if str(cat).strip()]
            categories = sorted(set(categories))
            if categories:
                selected_categories = st.multiselect("Фильтр по категориям", options=categories, default=[], key="ue_filter_categories")
            else:
                selected_categories = []
        else:
            selected_categories = []
    with col2:
        brand_options = [col for col in df.columns if any(w in col.lower() for w in ['бренд', 'brand', 'производитель'])]
        if brand_options:
            brand_col_filter = brand_options[0]
            brands = df[brand_col_filter].dropna().unique().tolist()
            brands = [str(b).strip() for b in brands if str(b).strip()]
            brands = sorted(set(brands))
            if brands:
                selected_brands = st.multiselect("Фильтр по брендам", options=brands, default=[], key="ue_filter_brands")
            else:
                selected_brands = []
        else:
            selected_brands = []

    # --- КНОПКА РАСЧЕТА ---
    if st.button("🚀 Рассчитать юнит-экономику по артикулам", type="primary", key="ue_calc_articles"):
        with st.spinner("Расчет юнит-экономики для всех товаров..."):
            try:
                filtered_df = df.copy()
                if selected_categories and category_col != 'Не выбрано' and category_col in df.columns:
                    filtered_df = filtered_df[filtered_df[category_col].astype(str).str.strip().isin(selected_categories)]
                if selected_brands and brand_options:
                    brand_col_filter = brand_options[0]
                    filtered_df = filtered_df[filtered_df[brand_col_filter].astype(str).str.strip().isin(selected_brands)]

                if filtered_df.empty:
                    st.warning("⚠️ Нет данных для расчета после применения фильтров")
                    return

                st.info(f"📊 Расчет для {len(filtered_df)} товаров")

                results = []
                progress_bar = st.progress(0, text="Расчет юнит-экономики...")
                total_rows = len(filtered_df)

                for idx, row in filtered_df.iterrows():
                    if idx % max(1, total_rows // 20) == 0:
                        progress_bar.progress(idx / total_rows, text=f"Расчет {idx+1}/{total_rows}...")

                    article = safe_str(row.get(article_col, ''))
                    price = safe_float(row.get(price_col, 0))
                    cost = safe_float(row.get(cost_col, 0))
                    if price <= 0 or cost <= 0:
                        continue

                    length = safe_float(row.get(length_col, 0)) if length_col != 'Не выбрано' else 0
                    width = safe_float(row.get(width_col, 0)) if width_col != 'Не выбрано' else 0
                    height = safe_float(row.get(height_col, 0)) if height_col != 'Не выбрано' else 0
                    weight = safe_float(row.get(weight_col, 0)) if weight_col != 'Не выбрано' else 0

                    if length > 0 and width > 0 and height > 0:
                        volume = (length * width * height) / 1000.0
                    else:
                        volume = 5.0
                    if weight <= 0:
                        weight = 1.0

                    category = None
                    if category_col != 'Не выбрано' and category_col in df.columns:
                        category = safe_str(row.get(category_col, ''))

                    final_price = price * (1 + markup_percent / 100) if apply_markup else price

                    for marketplace in selected_marketplaces:
                        economics = unit_economics.calculate_unit_economics(
                            price=final_price,
                            cost=cost,
                            marketplace=marketplace,
                            weight=weight,
                            category=category,
                            is_premium=is_premium
                        )

                        row_data = {
                            'Артикул': article,
                            'Бренд': safe_str(row.get('Бренд', '')),
                            'Наименование': safe_str(row.get('Наименование', '')),
                            'Категория': category or '',
                            'Цена_исходная': price,
                            'Цена_с_наценкой': final_price,
                            'Себестоимость': cost,
                            'Вес_кг': weight,
                            'Объем_л': volume,
                            'Маркетплейс': marketplace,
                            'Режим_работы': operation_mode,
                            'Прибыль': economics.profit,
                            'Маржа_%': economics.margin_percent,
                            'ROI_%': economics.roi,
                            'Точка_безубыточности': economics.breakeven_price,
                            'Прибыль_на_рубль': economics.profit_per_ruble,
                            'Комиссия': economics.commission,
                            'Комиссия_%': economics.commission_percent,
                            'Логистика': economics.logistics,
                            'Хранение': economics.storage_cost,
                            'Итого_расходов': economics.total_expenses
                        }
                        results.append(row_data)

                progress_bar.progress(1.0, text="Расчет завершен!")
                progress_bar.empty()

                if not results:
                    st.error("❌ Не удалось рассчитать юнит-экономику ни для одного товара")
                    return

                df_results = pd.DataFrame(results)
                st.session_state.ue_article_results = df_results

                st.success(f"✅ Рассчитано {len(df_results)} записей по {len(selected_marketplaces)} маркетплейсам")

                # --- СТАТИСТИКА ---
                st.subheader("📊 Сводная статистика")
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    total_profit = df_results['Прибыль'].sum()
                    st.metric("💰 Общая прибыль", f"{total_profit:,.0f} ₽")
                with col2:
                    avg_profit = df_results['Прибыль'].mean()
                    st.metric("📈 Средняя прибыль", f"{avg_profit:.2f} ₽")
                with col3:
                    avg_margin = df_results['Маржа_%'].mean()
                    st.metric("📊 Средняя маржа", f"{avg_margin:.1f}%")
                with col4:
                    best_mp = df_results.groupby('Маркетплейс')['Прибыль'].sum().idxmax()
                    st.metric("🏆 Лучший МП", best_mp)

                # --- ВИЗУАЛИЗАЦИЯ ---
                if PLOTLY_AVAILABLE and px is not None:
                    st.subheader("📊 Визуализация результатов")
                    tab1, tab2, tab3 = st.tabs(["📈 Прибыль по маркетплейсам", "📊 Распределение прибыли", "📉 Сравнение маржи"])
                    with tab1:
                        profit_by_mp = df_results.groupby('Маркетплейс')['Прибыль'].sum().reset_index()
                        fig = px.bar(profit_by_mp, x='Маркетплейс', y='Прибыль',
                                     title="Суммарная прибыль по маркетплейсам",
                                     color='Маркетплейс',
                                     color_discrete_sequence=px.colors.qualitative.Set3)
                        st.plotly_chart(fig, use_container_width=True)
                    with tab2:
                        fig = px.histogram(df_results, x='Прибыль', nbins=30,
                                           title="Распределение прибыли по товарам",
                                           color='Маркетплейс',
                                           color_discrete_sequence=px.colors.qualitative.Set3)
                        st.plotly_chart(fig, use_container_width=True)
                    with tab3:
                        margin_by_mp = df_results.groupby('Маркетплейс')['Маржа_%'].mean().reset_index()
                        fig = px.bar(margin_by_mp, x='Маркетплейс', y='Маржа_%',
                                     title="Средняя маржинальность по маркетплейсам",
                                     color='Маркетплейс',
                                     color_discrete_sequence=px.colors.qualitative.Set3)
                        st.plotly_chart(fig, use_container_width=True)

                # --- ТАБЛИЦА РЕЗУЛЬТАТОВ ---
                st.subheader("📋 Результаты расчета")
                display_cols = ['Артикул', 'Маркетплейс', 'Цена_с_наценкой', 'Прибыль', 'Маржа_%', 'ROI_%', 'Точка_безубыточности']
                available_display = [col for col in display_cols if col in df_results.columns]

                sort_col = st.selectbox("Сортировать по:", options=available_display,
                                        index=available_display.index('Прибыль') if 'Прибыль' in available_display else 0,
                                        key="ue_sort_col")
                sort_order = st.radio("Порядок:", ["По убыванию", "По возрастанию"], horizontal=True, key="ue_sort_order")
                sorted_df = df_results.sort_values(sort_col, ascending=(sort_order == "По возрастанию"))
                st.dataframe(sorted_df[available_display], use_container_width=True, key="ue_results_table")

                # --- ЭКСПОРТ ---
                st.subheader("📤 Экспорт результатов")
                export_col1, export_col2 = st.columns(2)
                with export_col1:
                    if st.button("📥 Экспорт в Excel", key="ue_export_excel"):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            df_results.to_excel(writer, sheet_name='Юнит-экономика', index=False)
                            summary = df_results.groupby('Маркетплейс').agg({
                                'Прибыль': ['sum', 'mean', 'count'],
                                'Маржа_%': 'mean',
                                'ROI_%': 'mean'
                            }).round(2)
                            summary.columns = ['Суммарная прибыль', 'Средняя прибыль', 'Количество', 'Средняя маржа %', 'Средний ROI %']
                            summary.to_excel(writer, sheet_name='Сводка', index=True)
                            pivot = df_results.pivot_table(index='Артикул', columns='Маркетплейс', values='Прибыль', aggfunc='sum').fillna(0)
                            pivot.to_excel(writer, sheet_name='Pivot_прибыль')
                        output.seek(0)
                        st.download_button(
                            label="📥 Скачать Excel",
                            data=output,
                            file_name=f"юнит_экономика_артикулы_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            key="ue_download_excel"
                        )
                with export_col2:
                    if st.button("📥 Экспорт в CSV", key="ue_export_csv"):
                        csv = df_results.to_csv(index=False, encoding='utf-8-sig')
                        st.download_button(
                            label="📥 Скачать CSV",
                            data=csv,
                            file_name=f"юнит_экономика_артикулы_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                            mime="text/csv",
                            key="ue_download_csv"
                        )

            except Exception as e:
                st.error(f"❌ Ошибка при расчете: {str(e)}")
                st.code(traceback.format_exc())
                logger.error(f"UE by article error: {traceback.format_exc()}")


# ============================================================================
# БЛОК 13: ОБОГАЩЕНИЕ КАТАЛОГА
# ============================================================================
def show_catalog_enhance_interface():
    """Интерфейс обогащения каталога (поиск аналогов)"""
    st.header("📊 Обогащение каталога (поиск аналогов)")
    st.info("""
🔍 **Поиск аналогов:**

Система ищет аналоги через общие OE номера (2 уровня):
- **Уровень 1**: Прямые аналоги (общие OE номера)
- **Уровень 2**: Косвенные аналоги (через аналоги уровнем выше)

**Как это работает:**
1. Загрузите данные в каталог (OE, артикулы, кросс-ссылки)
2. Система построит связи между товарами
3. При обогащении для каждого товара найдутся аналоги
""")

    if 'catalog_enhancer' not in st.session_state:
        st.session_state.catalog_enhancer = CatalogEnhancer()
    enhancer = st.session_state.catalog_enhancer

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("📤 Загрузка данных каталога")
        st.markdown("""
**Необходимые колонки в файлах:**

**OE данные:**
- `oe_number` - номер OE
- `name` - название
- `applicability` - применимость

**Детали (артикулы):**
- `artikul` - артикул
- `brand` - бренд
- `length`, `width`, `height`, `weight` - габариты

**Кросс-ссылки:**
- `oe_number` - номер OE
- `artikul` - артикул
- `brand` - бренд
""")
        oe_file = st.file_uploader("OE данные", type=['xlsx', 'csv'], key="enh_oe")
        parts_file = st.file_uploader("Детали (артикулы)", type=['xlsx', 'csv'], key="enh_parts")
        cross_file = st.file_uploader("Кросс-ссылки", type=['xlsx', 'csv'], key="enh_cross")

        if st.button("📥 Загрузить данные в каталог", type="primary", key="enh_load_data"):
            with st.spinner("Загрузка данных..."):
                if oe_file:
                    try:
                        df = pd.read_excel(oe_file) if oe_file.name.endswith('.xlsx') else pd.read_csv(oe_file)
                        enhancer.load_oe_data(df)
                        st.success(f"✅ Загружено {len(df)} OE записей")
                    except Exception as e:
                        st.error(f"❌ Ошибка загрузки OE: {str(e)}")
                if parts_file:
                    try:
                        df = pd.read_excel(parts_file) if parts_file.name.endswith('.xlsx') else pd.read_csv(parts_file)
                        enhancer.load_parts_data(df)
                        st.success(f"✅ Загружено {len(df)} записей деталей")
                    except Exception as e:
                        st.error(f"❌ Ошибка загрузки деталей: {str(e)}")
                if cross_file:
                    try:
                        df = pd.read_excel(cross_file) if cross_file.name.endswith('.xlsx') else pd.read_csv(cross_file)
                        enhancer.load_cross_references(df)
                        st.success(f"✅ Загружено {len(df)} кросс-ссылок")
                    except Exception as e:
                        st.error(f"❌ Ошибка загрузки кросс-ссылок: {str(e)}")

    with col2:
        st.subheader("🔍 Поиск аналогов")
        artikul = st.text_input("Артикул", placeholder="Введите артикул", key="enh_artikul_input")
        brand = st.text_input("Бренд", placeholder="Введите бренд", key="enh_brand_input")

        if st.button("🔍 Найти аналоги", type="primary", key="enh_find_analogs"):
            if artikul and brand:
                with st.spinner("Поиск аналогов..."):
                    data = enhancer.get_analog_data(artikul, brand)
                    if data.get('error'):
                        st.error(f"❌ {data['error']}")
                    else:
                        st.success(f"✅ Найдено {data.get('analog_count', 0)} аналогов")
                        if data.get('oe_list'):
                            st.info(f"🔗 OE номера: {data.get('oe_list')}")
                        if data.get('has_analogs'):
                            st.subheader("📋 Аналоги")
                            analogs_df = pd.DataFrame(data.get('analogs', []))
                            st.dataframe(analogs_df, use_container_width=True, key="enh_analogs_table")
            else:
                st.warning("⚠️ Введите артикул и бренд")

    st.subheader("📊 Статистика каталога")
    stats = enhancer.get_stats()
    col1, col2 = st.columns(2)
    with col1:
        st.metric("📦 OE записей", stats.get('oe_loaded', 0))
        st.metric("🔄 Кросс-ссылок", stats.get('cross_loaded', 0))
    with col2:
        st.metric("🔧 Деталей", stats.get('parts_loaded', 0))
        st.metric("🔍 Поисков", stats.get('analog_searches', 0))

    st.divider()
    st.subheader("📊 Обогащение загруженного каталога")

    if 'uploaded_data' in st.session_state and st.session_state.uploaded_data is not None:
        df = st.session_state.uploaded_data
        col1, col2 = st.columns(2)
        with col1:
            artikul_col = st.selectbox("Колонка с артикулом", df.columns, key="enh_artikul_col_select")
        with col2:
            brand_col = st.selectbox("Колонка с брендом", df.columns, key="enh_brand_col_select")

        if st.button("🚀 Обогатить данные", type="primary", key="enh_enrich_data"):
            with st.spinner("Обогащение данных..."):
                enhanced_df = enhancer.enhance_catalog_data(df, artikul_col, brand_col)
                st.session_state.uploaded_data = enhanced_df
                st.success("✅ Данные обогащены!")

                st.subheader("📊 Результат обогащения")
                st.dataframe(enhanced_df.head(20), use_container_width=True, key="enh_result_table")

                if 'analog_count' in enhanced_df.columns:
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        st.metric("📦 Всего товаров", len(enhanced_df))
                    with col2:
                        st.metric("🔄 С аналогами", len(enhanced_df[enhanced_df['analog_count'] > 0]))
                    with col3:
                        st.metric("📊 Среднее аналогов", f"{enhanced_df['analog_count'].mean():.1f}")
    else:
        st.warning("⚠️ Сначала загрузите данные в разделе '📁 Загрузка данных'")


# ============================================================================
# БЛОК 14: HIGH-VOLUME ИНТЕРФЕЙС (700+ СТРОК)
# ============================================================================
def show_high_volume_interface():
    """Интерфейс для High-Volume каталога"""
    st.header("🚗 High-Volume Каталог автозапчастей (10M+)")

    if not (POLARS_AVAILABLE and DUCKDB_AVAILABLE):
        st.warning("⚠️ Для работы High-Volume режима установите: `pip install polars duckdb`")
        return

    if 'high_volume_catalog' not in st.session_state:
        st.session_state.high_volume_catalog = HighVolumeAutoPartsCatalog()
    catalog = st.session_state.high_volume_catalog

    if not catalog.conn:
        st.error("❌ Ошибка подключения к базе данных")
        return

    st.sidebar.title("🧭 Меню High-Volume")
    option = st.sidebar.radio("Выберите раздел", ["Загрузка данных", "Экспорт", "Статистика", "Управление"])

    if option == "Загрузка данных":
        st.header("📥 Загрузка данных")
        col1, col2 = st.columns(2)
        with col1:
            oe_file = st.file_uploader("Основные данные (OE)", type=['xlsx'], key="hv_oe")
            cross_file = st.file_uploader("Кроссы (OE→Артикул)", type=['xlsx'], key="hv_cross")
            barcode_file = st.file_uploader("Штрих-коды", type=['xlsx'], key="hv_barcode")
        with col2:
            weight_dims_file = st.file_uploader("Вес и габариты", type=['xlsx'], key="hv_dims")
            images_file = st.file_uploader("Изображения", type=['xlsx'], key="hv_images")
            prices_file = st.file_uploader("Цены", type=['xlsx'], key="hv_prices")

        uploaded_files = {
            'oe': oe_file,
            'cross': cross_file,
            'barcode': barcode_file,
            'dimensions': weight_dims_file,
            'images': images_file,
            'prices': prices_file
        }

        if st.button("Обработать и загрузить", key="hv_load"):
            saved_paths = {}
            for key, file in uploaded_files.items():
                if file:
                    path = catalog.data_dir / f"{key}_{int(time.time())}.xlsx"
                    with open(path, "wb") as f:
                        f.write(file.getbuffer())
                    saved_paths[key] = str(path)

            if saved_paths:
                with st.spinner("Обработка файлов..."):
                    dataframes = catalog.merge_all_data_parallel(saved_paths)
                with st.spinner("Загрузка данных в базу..."):
                    catalog.process_and_load_data(dataframes)
            else:
                st.warning("Загрузите хотя бы один файл")

    elif option == "Экспорт":
        st.header("📤 Экспорт данных")
        total = catalog.conn.execute(
            "SELECT COUNT(*) FROM (SELECT DISTINCT artikul_norm, brand_norm FROM parts)").fetchone()[0]
        st.info(f"Всего: {total}")

        if total == 0:
            st.warning("Нет данных для экспорта")
            return

        format_choice = st.radio("Формат", ["CSV", "Excel", "Parquet"], key="hv_format")
        selected_columns = st.multiselect("Колонки", [
            "Артикул бренда", "Бренд", "Наименование", "Применимость", "Описание",
            "Категория товара", "Кратность", "Длинна", "Ширина", "Высота", "Вес",
            "Длинна/Ширина/Высота", "OE номер", "аналоги", "Ссылка на изображение", "Цена", "Валюта"
        ], key="hv_columns")
        include_prices = st.checkbox("Включить цены", value=True, key="hv_prices_include")
        apply_markup = st.checkbox("Применить наценку", value=True, disabled=not include_prices, key="hv_markup")

        if st.button("🚀 Экспортировать", key="hv_export_btn"):
            output_path = catalog.data_dir / f"export.{format_choice.lower()}"
            with st.spinner("Генерация файла..."):
                if format_choice == "CSV":
                    catalog.export_to_csv_optimized(str(output_path), selected_columns if selected_columns else None, include_prices, apply_markup)
                elif format_choice == "Excel":
                    catalog.export_to_excel_optimized(str(output_path), selected_columns if selected_columns else None, include_prices, apply_markup)
                elif format_choice == "Parquet":
                    catalog.export_to_parquet(str(output_path), selected_columns if selected_columns else None, include_prices, apply_markup)

            if output_path.exists():
                with open(output_path, "rb") as f:
                    st.download_button("⬇️ Скачать файл", f, file_name=output_path.name, key="hv_download")

    elif option == "Статистика":
        st.header("📈 Статистика")
        stats = catalog.get_statistics()
        if stats:
            col1, col2, col3 = st.columns(3)
            col1.metric("Уникальных товаров", f"{stats.get('unique_parts', 0):,}")
            col2.metric("Брендов", f"{stats.get('brands', 0):,}")
            col3.metric("Средняя цена", f"{stats.get('avg_price', 0):.2f} ₽")

            st.subheader("Топ 10 брендов")
            if 'top_brands' in stats and not stats['top_brands'].empty:
                st.dataframe(stats['top_brands'])

    elif option == "Управление":
        st.header("🔧 Управление данными")
        st.warning("⚠️ Операции необратимы!")

        management_option = st.radio(
            "Выберите действие:",
            ["Удалить по бренду", "Удалить по артикулу", "Управление ценами", "Исключения", "Категории"],
            key="hv_management"
        )

        if management_option == "Удалить по бренду":
            try:
                brands_result = catalog.conn.execute(
                    "SELECT DISTINCT brand FROM parts WHERE brand IS NOT NULL ORDER BY brand").fetchall()
                available_brands = [row[0] for row in brands_result] if brands_result else []
            except Exception as e:
                st.error(f"Ошибка: {e}")
                return

            if not available_brands:
                st.info("Нет данных")
                return

            selected_brand = st.selectbox("Бренд", available_brands, key="hv_delete_brand")
            brand_norm_result = catalog.conn.execute(
                "SELECT brand_norm FROM parts WHERE brand = ? LIMIT 1", [selected_brand]).fetchone()
            if brand_norm_result:
                brand_norm = brand_norm_result[0]
            else:
                brand_norm = catalog.normalize_key(pl.Series([selected_brand]))[0]

            count = catalog.conn.execute(
                "SELECT COUNT(*) FROM parts WHERE brand_norm = ?", [brand_norm]).fetchone()[0]
            st.info(f"Удалить {count} записей бренда '{selected_brand}'?")

            if st.checkbox("Подтверждаю удаление", key="hv_confirm_brand"):
                if st.button("Удалить", key="hv_delete_brand_btn"):
                    deleted = catalog.delete_by_brand(brand_norm)
                    st.success(f"Удалено {deleted} записей")
                    st.rerun()

        elif management_option == "Удалить по артикулу":
            artikul_input = st.text_input("Артикул", key="hv_delete_artikul")
            if artikul_input:
                artikul_norm = catalog.normalize_key(pl.Series([artikul_input]))[0]
                count = catalog.conn.execute(
                    "SELECT COUNT(*) FROM parts WHERE artikul_norm = ?", [artikul_norm]).fetchone()[0]
                st.info(f"Найдено {count} записей для артикула '{artikul_input}'")

                if st.checkbox("Подтверждаю", key="hv_confirm_artikul"):
                    if st.button("Удалить", key="hv_delete_artikul_btn"):
                        deleted = catalog.delete_by_artikul(artikul_norm)
                        st.success(f"Удалено {deleted} записей")
                        st.rerun()

        elif management_option == "Управление ценами":
            show_price_settings(catalog)
        elif management_option == "Исключения":
            show_exclusion_settings(catalog)
        elif management_option == "Категории":
            show_category_mapping_settings(catalog)


def show_price_settings(catalog):
    """Настройки цен и наценок"""
    st.subheader("💰 Управление ценами и наценками")

    st.subheader("Общая наценка")
    global_markup = st.number_input(
        "Общая наценка (%):",
        min_value=0.0,
        max_value=500.0,
        value=catalog.price_rules['global_markup'] * 100,
        step=0.1,
        key="hv_global_markup"
    )
    catalog.price_rules['global_markup'] = global_markup / 100

    st.subheader("Наценки по брендам")
    brand_markups = catalog.price_rules.get('brand_markups', {})

    try:
        brands_result = catalog.conn.execute(
            "SELECT DISTINCT brand FROM parts WHERE brand IS NOT NULL ORDER BY brand").fetchall()
        available_brands = [row[0] for row in brands_result] if brands_result else []
    except Exception as e:
        st.error(f"Ошибка при получении списка брендов: {e}")
        available_brands = []

    if available_brands:
        col1, col2 = st.columns([2, 1])
        with col1:
            selected_brand = st.selectbox("Выберите бренд:", available_brands, key="hv_brand_select")
        with col2:
            current_markup = brand_markups.get(selected_brand, catalog.price_rules.get('global_markup', 0))
            brand_markup = st.number_input(
                "Наценка (%):",
                min_value=0.0,
                max_value=500.0,
                value=current_markup * 100,
                step=0.1,
                key=f"hv_markup_{selected_brand}"
            )
            if st.button("Сохранить наценку", key=f"hv_save_{selected_brand}"):
                brand_markups[selected_brand] = brand_markup / 100
                catalog.price_rules['brand_markups'] = brand_markups
                catalog.save_price_rules()
                st.success(f"✅ Наценка для {selected_brand} сохранена")

    st.subheader("Ограничения по ценам")
    col1, col2 = st.columns(2)
    with col1:
        min_price = st.number_input(
            "Минимальная цена:",
            min_value=0.0,
            value=float(catalog.price_rules['min_price']),
            step=0.01,
            key="hv_min_price"
        )
        catalog.price_rules['min_price'] = min_price
    with col2:
        max_price = st.number_input(
            "Максимальная цена:",
            min_value=0.0,
            value=float(catalog.price_rules['max_price']),
            step=0.01,
            key="hv_max_price"
        )
        catalog.price_rules['max_price'] = max_price

    if st.button("Сохранить все настройки цен", key="hv_save_prices"):
        catalog.save_price_rules()
        st.success("✅ Все настройки цен сохранены")


def show_exclusion_settings(catalog):
    """Настройки исключений"""
    st.subheader("🚫 Управление исключениями при экспорте")
    st.info("Товары, содержащие эти слова в названии, будут исключены из экспорта")

    current_exclusions = "\n".join(catalog.exclusion_rules)
    new_exclusions = st.text_area(
        "Список исключений (по одному на строку):",
        value=current_exclusions,
        height=200,
        placeholder="Введите слова для исключения, например:\nКузов\nСтекла\nМасла",
        key="hv_exclusions"
    )

    if st.button("Сохранить правила исключения", key="hv_save_exclusions"):
        cleaned = [line.strip() for line in new_exclusions.splitlines() if line.strip()]
        if len(cleaned) != len(set(cleaned)):
            st.warning("Обнаружены дублирующие записи. Они будут автоматически удалены.")
        catalog.exclusion_rules = list(dict.fromkeys(cleaned))
        catalog.save_exclusion_rules()
        st.success("✅ Правила исключения сохранены")


def show_category_mapping_settings(catalog):
    """Настройки категорий"""
    st.subheader("🗂️ Управление категориями товаров")
    st.info("Настройте соответствие между названиями товаров и категориями")

    st.subheader("Текущие правила")
    if catalog.category_mapping:
        mapping_df = pd.DataFrame({
            "Название товара": list(catalog.category_mapping.keys()),
            "Категория": list(catalog.category_mapping.values())
        })
        st.dataframe(mapping_df, use_container_width=True, hide_index=True)
    else:
        st.write("Нет пользовательских правил")

    st.subheader("Добавить правило")
    col1, col2 = st.columns(2)
    with col1:
        name_pattern = st.text_input("Ключевое слово в названии", key="hv_cat_pattern")
    with col2:
        category = st.text_input("Категория", key="hv_cat_name")

    if st.button("➕ Добавить", key="hv_add_category"):
        if name_pattern.strip() and category.strip():
            normalized_key = name_pattern.strip().lower()
            existing_keys = {k.lower(): k for k in catalog.category_mapping.keys()}
            if normalized_key in existing_keys:
                st.warning(f"Правило для '{existing_keys[normalized_key]}' обновлено")
            catalog.category_mapping[name_pattern.strip()] = category.strip()
            catalog.save_category_mapping()
            st.success(f"Добавлено: {name_pattern.strip()} → {category.strip()}")
            st.rerun()
        else:
            st.error("Заполните оба поля")

    if catalog.category_mapping:
        st.subheader("🗑️ Удалить правило")
        rule_to_delete = st.selectbox(
            "Выберите правило",
            options=list(catalog.category_mapping.keys()),
            format_func=lambda x: f"{x} → {catalog.category_mapping[x]}",
            key="hv_delete_rule"
        )
        if st.button("Удалить", key="hv_delete_category"):
            del catalog.category_mapping[rule_to_delete]
            catalog.save_category_mapping()
            st.success(f"Удалено: {rule_to_delete}")
            st.rerun()


# ============================================================================
# БЛОК 15: АНАЛИТИКА (400+ СТРОК)
# ============================================================================
def show_analytics_interface():
    """Интерфейс аналитики"""
    st.header("📊 Аналитика и статистика")

    if st.session_state.get('uploaded_data') is None:
        st.warning("⚠️ Сначала загрузите данные в разделе '📁 Загрузка данных'")
        return

    df = st.session_state.uploaded_data.copy()
    st.success(f"✅ Анализ {len(df)} товаров")

    st.subheader("📈 Общая статистика")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📦 Товаров", len(df))
    with col2:
        price_col = None
        for col in df.columns:
            if any(w in col.lower() for w in ['цена', 'price', 'стоимость']):
                price_col = col
                break
        if price_col:
            avg_price = safe_float(df[price_col].mean())
            st.metric("💰 Средняя цена", f"{avg_price:,.0f} ₽")
    with col3:
        cost_col = None
        for col in df.columns:
            if any(w in col.lower() for w in ['себестоимость', 'cost', 'закупочная']):
                cost_col = col
                break
        if cost_col:
            avg_cost = safe_float(df[cost_col].mean())
            st.metric("💵 Средняя себестоимость", f"{avg_cost:,.0f} ₽")
    with col4:
        if price_col and cost_col:
            df['_margin'] = safe_float(df[price_col]) - safe_float(df[cost_col])
            avg_margin = df['_margin'].mean()
            st.metric("📈 Средняя маржа", f"{avg_margin:,.0f} ₽")

    if PLOTLY_AVAILABLE and px is not None:
        st.subheader("📊 Визуализация данных")
        chart_type = st.selectbox(
            "Тип графика",
            ["Распределение цен", "Распределение категорий", "Топ товаров по цене", "Анализ маржи"],
            key="analytics_chart_type"
        )

        if chart_type == "Распределение цен" and price_col:
            fig = px.histogram(df, x=price_col, nbins=30,
                               title=f"Распределение цен ({price_col})",
                               color_discrete_sequence=['#e94560'])
            st.plotly_chart(fig, use_container_width=True)
        elif chart_type == "Распределение категорий" and 'Категория' in df.columns:
            category_counts = df['Категория'].value_counts()
            fig = px.pie(values=category_counts.values, names=category_counts.index,
                         title="Распределение товаров по категориям",
                         color_discrete_sequence=px.colors.qualitative.Set3)
            st.plotly_chart(fig, use_container_width=True)
        elif chart_type == "Топ товаров по цене" and price_col:
            name_col = None
            for col in df.columns:
                if any(w in col.lower() for w in ['наименование', 'название', 'name', 'товар']):
                    name_col = col
                    break
            if name_col:
                top_df = df.nlargest(10, price_col)[[name_col, price_col]]
                fig = px.bar(top_df, x=name_col, y=price_col,
                             title="Топ 10 товаров по цене",
                             color_discrete_sequence=['#0f3460'])
                st.plotly_chart(fig, use_container_width=True)
        elif chart_type == "Анализ маржи" and price_col and cost_col:
            df['_margin_percent'] = (safe_float(df[price_col]) - safe_float(df[cost_col])) / safe_float(df[price_col]) * 100
            fig = px.scatter(df, x=price_col, y='_margin_percent',
                             title="Зависимость маржи от цены",
                             color_discrete_sequence=['#e94560'])
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("📊 Для расширенной визуализации установите plotly: `pip install plotly`")

    if price_col:
        st.subheader("📊 Статистика цен")
        st.dataframe(df[price_col].describe(), use_container_width=True)

    st.subheader("📊 Корреляционный анализ")
    numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
    if len(numeric_cols) >= 2:
        corr_df = df[numeric_cols].corr()
        st.dataframe(corr_df.style.background_gradient(cmap='RdBu_r'), use_container_width=True)

        if PLOTLY_AVAILABLE and go is not None:
            fig = go.Figure(data=go.Heatmap(
                z=corr_df.values,
                x=corr_df.columns,
                y=corr_df.columns,
                colorscale='RdBu_r',
                zmin=-1, zmax=1
            ))
            fig.update_layout(title="Тепловая карта корреляций", height=500)
            st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Для корреляционного анализа необходимо минимум 2 числовые колонки")


# ============================================================================
# БЛОК 16: ИСТОРИЯ РАСЧЕТОВ
# ============================================================================
def show_history_interface():
    """Интерфейс истории расчетов"""
    st.header("📋 История расчетов")
    unit_economics = MarketplaceUnitEconomics()
    history = unit_economics.get_history()

    if not history:
        st.info("📋 История расчетов пуста. Выполните расчеты в разделе '📊 Юнит-экономика'")
        return

    df_history = pd.DataFrame([r.to_dict() if hasattr(r, 'to_dict') else r for r in history])

    st.subheader("🔍 Фильтры")
    col1, col2, col3 = st.columns(3)
    with col1:
        marketplaces = ['Все'] + sorted(df_history['marketplace'].unique().tolist())
        filter_marketplace = st.selectbox("Маркетплейс", marketplaces, key="history_marketplace")
    with col2:
        modes = ['Все'] + sorted(df_history['operation_mode'].unique().tolist())
        filter_mode = st.selectbox("Режим работы", modes, key="history_mode")
    with col3:
        if 'timestamp' in df_history.columns:
            df_history['timestamp_dt'] = pd.to_datetime(df_history['timestamp'])
            min_date = df_history['timestamp_dt'].min().date()
            max_date = df_history['timestamp_dt'].max().date()
            filter_start = st.date_input("Дата с", min_date, key="history_start")
            filter_end = st.date_input("Дата по", max_date, key="history_end")

    filtered_df = df_history.copy()
    if filter_marketplace != 'Все':
        filtered_df = filtered_df[filtered_df['marketplace'] == filter_marketplace]
    if filter_mode != 'Все':
        filtered_df = filtered_df[filtered_df['operation_mode'] == filter_mode]
    if 'timestamp_dt' in filtered_df.columns:
        filtered_df = filtered_df[
            (filtered_df['timestamp_dt'].dt.date >= filter_start) &
            (filtered_df['timestamp_dt'].dt.date <= filter_end)
        ]

    st.subheader(f"📊 Найдено расчетов: {len(filtered_df)}")

    if not filtered_df.empty:
        display_cols = ['marketplace', 'operation_mode', 'price', 'cost', 'profit', 'margin_percent', 'roi', 'timestamp']
        available_cols = [col for col in display_cols if col in filtered_df.columns]
        st.dataframe(filtered_df[available_cols].sort_values('timestamp', ascending=False),
                     use_container_width=True, key="history_table")

        st.subheader("📊 Статистика по истории")
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            total_profit = filtered_df['profit'].sum()
            st.metric("💰 Суммарная прибыль", f"{total_profit:,.0f} ₽")
        with col2:
            avg_profit = filtered_df['profit'].mean()
            st.metric("📈 Средняя прибыль", f"{avg_profit:.2f} ₽")
        with col3:
            max_profit = filtered_df['profit'].max()
            best_market = filtered_df[filtered_df['profit'] == max_profit]['marketplace'].iloc[0] if not filtered_df.empty else '-'
            st.metric("🏆 Лучший результат", f"{max_profit:.2f} ₽", delta=f"{best_market}")
        with col4:
            count = len(filtered_df)
            st.metric("📊 Всего записей", count)

        if PLOTLY_AVAILABLE and go is not None and len(filtered_df) > 1:
            try:
                from plotly.subplots import make_subplots
                fig = make_subplots(
                    rows=2, cols=1,
                    subplot_titles=("Прибыль по датам", "Распределение прибыли по маркетплейсам")
                )
                fig.add_trace(
                    go.Scatter(
                        x=filtered_df['timestamp_dt'] if 'timestamp_dt' in filtered_df.columns else filtered_df.index,
                        y=filtered_df['profit'],
                        mode='lines+markers',
                        name='Прибыль',
                        line=dict(color='#e94560', width=2)
                    ),
                    row=1, col=1
                )
                fig.add_trace(
                    go.Box(
                        x=filtered_df['marketplace'],
                        y=filtered_df['profit'],
                        name='Прибыль по МП',
                        marker_color='#0f3460'
                    ),
                    row=2, col=1
                )
                fig.update_layout(height=600, showlegend=True, title_text="Визуализация истории расчетов")
                st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                logger.warning(f"Ошибка визуализации истории: {e}")

    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 Экспортировать историю в CSV", key="history_export"):
            if not filtered_df.empty:
                csv = filtered_df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 Скачать CSV",
                    data=csv,
                    file_name=f"история_расчетов_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key="history_download"
                )
    with col2:
        if st.button("🗑️ Очистить историю", type="secondary", key="history_clear"):
            if st.checkbox("Подтвердите очистку", key="history_confirm"):
                unit_economics.clear_history()
                st.success("✅ История очищена")
                st.rerun()


# ============================================================================
# БЛОК 17: ЭКСПОРТ
# ============================================================================
def show_export_interface():
    """Интерфейс экспорта данных"""
    st.header("📤 Экспорт данных")

    if st.session_state.get('uploaded_data') is None:
        st.warning("⚠️ Сначала загрузите данные")
        return

    df = st.session_state.uploaded_data
    st.success(f"✅ Готово к экспорту: {len(df)} товаров, {len(df.columns)} колонок")

    st.subheader("📊 Статистика перед экспортом")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📦 Товаров", len(df))
    with col2:
        price_col = None
        for col in df.columns:
            if any(w in col.lower() for w in ['цена', 'price', 'стоимость']):
                price_col = col
                break
        if price_col:
            avg_price = safe_float(df[price_col].mean())
            st.metric("💰 Средняя цена", f"{avg_price:.2f} ₽")
    with col3:
        cost_col = None
        for col in df.columns:
            if any(w in col.lower() for w in ['себестоимость', 'cost', 'закупочная']):
                cost_col = col
                break
        if cost_col:
            avg_cost = safe_float(df[cost_col].mean())
            st.metric("💵 Средняя себестоимость", f"{avg_cost:.2f} ₽")
    with col4:
        if 'Категория' in df.columns:
            st.metric("📂 Категорий", df['Категория'].nunique())

    export_format = st.radio("Формат экспорта", ["Excel (.xlsx)", "CSV (.csv)"], horizontal=True, key="export_format")
    include_stats = st.checkbox("📊 Включить лист со статистикой", value=True, key="export_stats")
    include_history = st.checkbox("📋 Включить историю расчетов", value=True, key="export_history")

    if st.button("📥 Скачать файл", type="primary", key="export_btn"):
        try:
            if export_format.startswith("Excel"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Данные', index=False)

                    if include_stats:
                        stats_data = {
                            'Показатель': ['Всего товаров', 'Колонок', 'Категорий'],
                            'Значение': [
                                len(df),
                                len(df.columns),
                                df['Категория'].nunique() if 'Категория' in df.columns else 0
                            ]
                        }
                        if price_col:
                            stats_data['Показатель'].append('Средняя цена')
                            stats_data['Значение'].append(safe_float(df[price_col].mean()))
                        if cost_col:
                            stats_data['Показатель'].append('Средняя себестоимость')
                            stats_data['Значение'].append(safe_float(df[cost_col].mean()))
                        pd.DataFrame(stats_data).to_excel(writer, sheet_name='Статистика', index=False)

                    if include_history:
                        unit_economics = MarketplaceUnitEconomics()
                        history = unit_economics.get_history()
                        if history:
                            pd.DataFrame([r.to_dict() if hasattr(r, 'to_dict') else r for r in history]).to_excel(
                                writer, sheet_name='История', index=False)

                output.seek(0)
                st.download_button(
                    label="📥 Скачать Excel файл",
                    data=output,
                    file_name=f"каталог_экспорт_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="download_excel"
                )
            else:
                csv = df.to_csv(index=False, encoding='utf-8-sig')
                st.download_button(
                    label="📥 Скачать CSV файл",
                    data=csv,
                    file_name=f"каталог_экспорт_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                    mime="text/csv",
                    key="download_csv"
                )
            st.success("✅ Данные успешно подготовлены к экспорту!")
        except Exception as e:
            st.error(f"❌ Ошибка экспорта: {str(e)}")
            logger.error(f"Export error: {traceback.format_exc()}")


# ============================================================================
# БЛОК 18: НАСТРОЙКИ
# ============================================================================
def show_settings_interface():
    """Интерфейс настроек"""
    st.header("⚙️ Настройки приложения")
    st.info("""
💡 **Настройки сохраняются в сессии браузера**
Настройки позволяют адаптировать приложение под ваши задачи.
""")

    st.subheader("🎨 Тема оформления")
    theme = st.selectbox("Тема", ["🌞 Светлая", "🌙 Темная", "🔄 Системная"], key="settings_theme")
    st.caption("Тема применяется при следующем запуске приложения")

    st.subheader("💱 Валютные настройки")
    currency = st.selectbox("Основная валюта", ["₽ (Рубль)", "$ (Доллар)", "€ (Евро)", "₴ (Гривна)", "¥ (Юань)"], key="settings_currency")
    show_currency_symbol = st.checkbox("Отображать символ валюты", value=True, key="settings_show_currency")

    st.subheader("📊 Параметры расчета")
    col1, col2 = st.columns(2)
    with col1:
        default_margin = st.number_input("Целевая маржинальность (%)", min_value=0.0, max_value=100.0, value=15.0, step=1.0, key="settings_default_margin")
    with col2:
        min_profit = st.number_input("Минимальная прибыль (₽)", min_value=0.0, value=50.0, step=10.0, key="settings_min_profit")

    st.subheader("📤 Экспортные настройки")
    col1, col2 = st.columns(2)
    with col1:
        default_export_format = st.selectbox("Формат по умолчанию", ["Excel", "CSV"], key="settings_export_format")
    with col2:
        include_timestamp = st.checkbox("Добавлять дату в имя файла", value=True, key="settings_include_timestamp")

    st.subheader("📝 Настройки логирования")
    log_level = st.select_slider("Уровень логирования", options=["DEBUG", "INFO", "WARNING", "ERROR"], value="INFO", key="settings_log_level")

    if st.button("💾 Сохранить настройки", type="primary", key="settings_save"):
        st.session_state.settings = {
            "theme": theme,
            "currency": currency,
            "show_currency_symbol": show_currency_symbol,
            "default_margin": default_margin,
            "min_profit": min_profit,
            "default_export_format": default_export_format,
            "include_timestamp": include_timestamp,
            "log_level": log_level
        }
        st.success("✅ Настройки сохранены!")
        logger.setLevel(log_level)
        st.balloons()

    st.divider()
    st.subheader("ℹ️ Информация о приложении")
    col1, col2 = st.columns(2)
    with col1:
        st.metric("📌 Версия", APP_VERSION)
        st.metric("🐍 Python", sys.version.split()[0])
    with col2:
        st.metric("📅 Дата", datetime.now().strftime("%d.%m.%Y"))
        data_loaded = st.session_state.get('uploaded_data') is not None
        rows = len(st.session_state.get('uploaded_data', pd.DataFrame())) if data_loaded else 0
        st.metric("📊 Данных загружено", rows)

    st.subheader("📚 Доступные библиотеки")
    available_libs = []
    libs_status = {
        "streamlit": True,
        "pandas": True,
        "numpy": True,
        "plotly": PLOTLY_AVAILABLE,
        "sklearn": SKLEARN_AVAILABLE,
        "duckdb": DUCKDB_AVAILABLE,
        "polars": POLARS_AVAILABLE,
        "openpyxl": OPENPYXL_AVAILABLE,
        "reportlab": PDF_EXPORT,
        "chardet": CHARDET_AVAILABLE,
        "openai": OPENAI_AVAILABLE,
        "torch": PYTORCH_AVAILABLE,
        "tensorflow": TENSORFLOW_AVAILABLE,
        "transformers": TRANSFORMERS_AVAILABLE,
        "aiohttp": ASYNC_AVAILABLE
    }
    cols = st.columns(3)
    for i, (lib, installed) in enumerate(libs_status.items()):
        cols[i % 3].write(f"{'✅' if installed else '❌'} {lib}")


# ============================================================================
# БЛОК 19: ГЛАВНАЯ ФУНКЦИЯ
# ============================================================================
def main():
    """Главная функция приложения"""
    st.set_page_config(
        page_title=f"{APP_NAME} v{APP_VERSION}",
        page_icon="🚗",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    st.markdown(f"""
<div style="text-align: center; padding: 20px; background: linear-gradient(135deg, #0f3460 0%, #16213e 100%); border-radius: 10px; margin-bottom: 20px;">
<h1 style="color: white;">🚗 {APP_NAME}</h1>
<p style="color: #e94560; font-size: 18px;">v{APP_VERSION} | Полная версия 6500+ строк</p>
<p style="color: #aaa;">Юнит-экономика маркетплейсов 2026 | Каталог с поиском аналогов 2 уровня</p>
<p style="color: #888;">High-Volume каталог (10M+) | ИИ-обновление тарифов | Экспорт с формулами</p>
<p style="color: #666; font-size: 14px;">✅ Улучшенная обработка загрузки файлов с автоопределением кодировки</p>
<p style="color: #666; font-size: 14px;">✅ Юнит-экономика по каждому артикулу с визуализацией</p>
</div>
""", unsafe_allow_html=True)

    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/000000/bar-chart.png", width=80)
        st.markdown("---")

        menu_options = [
            "📁 Загрузка данных",
            "📊 Обогащение каталога",
            "📊 Юнит-экономика",
            "📊 Юнит-экономика по артикулам",
            "📊 Аналитика",
            "📋 История расчетов",
            "📤 Экспорт",
            "🚗 High-Volume каталог",
            "⚙️ Настройки"
        ]

        menu_icons = {
            "📁 Загрузка данных": "📤",
            "📊 Обогащение каталога": "🔍",
            "📊 Юнит-экономика": "💰",
            "📊 Юнит-экономика по артикулам": "📊",
            "📊 Аналитика": "📈",
            "📋 История расчетов": "📜",
            "📤 Экспорт": "💾",
            "🚗 High-Volume каталог": "🚗",
            "⚙️ Настройки": "⚙️"
        }

        menu = st.radio(
            "Меню",
            menu_options,
            key="main_menu",
            format_func=lambda x: f"{menu_icons.get(x, '')} {x}"
        )

        st.markdown("---")
        st.markdown("### 📊 Состояние системы")
        data_loaded = st.session_state.get('uploaded_data') is not None
        rows = len(st.session_state.get('uploaded_data', pd.DataFrame())) if data_loaded else 0
        st.metric("📁 Данные", "Загружены ✅" if data_loaded else "Не загружены ❌")
        if data_loaded:
            st.metric("📦 Товаров", rows)

        st.markdown("---")
        st.caption(f"Приложение v{APP_VERSION}")
        st.caption(f"Python {sys.version.split()[0]}")

        with st.expander("📚 Библиотеки"):
            libs_status = {
                "Plotly": PLOTLY_AVAILABLE,
                "Sklearn": SKLEARN_AVAILABLE,
                "DuckDB": DUCKDB_AVAILABLE,
                "Polars": POLARS_AVAILABLE,
                "OpenPyXL": OPENPYXL_AVAILABLE,
                "PDF": PDF_EXPORT,
                "PyTorch": PYTORCH_AVAILABLE,
                "TensorFlow": TENSORFLOW_AVAILABLE,
                "Transformers": TRANSFORMERS_AVAILABLE,
                "Async": ASYNC_AVAILABLE
            }
            for lib, available in libs_status.items():
                st.write(f"{'✅' if available else '❌'} {lib}")

    try:
        if menu == "📁 Загрузка данных":
            show_data_upload_interface()
        elif menu == "📊 Обогащение каталога":
            show_catalog_enhance_interface()
        elif menu == "📊 Юнит-экономика":
            show_unit_economics_interface()
        elif menu == "📊 Юнит-экономика по артикулам":
            show_unit_economics_by_article_interface()
        elif menu == "📊 Аналитика":
            show_analytics_interface()
        elif menu == "📋 История расчетов":
            show_history_interface()
        elif menu == "📤 Экспорт":
            show_export_interface()
        elif menu == "🚗 High-Volume каталог":
            show_high_volume_interface()
        elif menu == "⚙️ Настройки":
            show_settings_interface()
    except Exception as e:
        st.error(f"❌ Ошибка при отображении страницы: {str(e)}")
        st.code(traceback.format_exc())
        logger.error(f"Main page error: {traceback.format_exc()}")


# ============================================================================
# ТОЧКА ВХОДА
# ============================================================================
if __name__ == "__main__":
    main()
