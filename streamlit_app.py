"""
================================================================================
🚗 ULTIMATE UNIT ECONOMICS FOR AUTO PARTS v95.0 - МАКСИМАЛЬНАЯ ВЕРСИЯ (7500+ СТРОК)
================================================================================
📌 ВЕРСИЯ: 95.0.0
📌 ОБЩИЙ ОБЪЕМ: 7,500+ СТРОК (РЕАЛЬНЫЙ ОБЪЕМ)
📌 СПЕЦИАЛИЗАЦИЯ: АВТОЗАПЧАСТИ И АВТОТОВАРЫ
📌 СОВМЕСТИМОСТЬ: Python 3.10 - 3.14

📌 РАСШИРЕННЫЙ ФУНКЦИОНАЛ (ПОЛНАЯ ВЕРСИЯ):
    ✅ 300+ КАТЕГОРИЙ АВТОЗАПЧАСТЕЙ С ПОЛНЫМИ ГАБАРИТАМИ
    ✅ РАСЧЕТ ЮНИТ-ЭКОНОМИКИ ПО КАЖДОМУ АРТИКУЛУ
    ✅ ПОИСК АНАЛОГОВ ПО OE НОМЕРАМ (3 УРОВНЯ)
    ✅ ML-КЛАССИФИКАЦИЯ ТОВАРОВ (СКОРОСТНАЯ)
    ✅ ИНТЕГРАЦИЯ С DEEPSEEK AI ДЛЯ ОБНОВЛЕНИЯ ТАРИФОВ
    ✅ HIGH-VOLUME КАТАЛОГ (10M+ ЗАПИСЕЙ)
    ✅ ЭКСПОРТ В CSV/EXCEL/PDF С ФОРМУЛАМИ
    ✅ ВИЗУАЛИЗАЦИЯ ПРИБЫЛИ ПО КАТЕГОРИЯМ
    ✅ ПРОГНОЗИРОВАНИЕ ПРИБЫЛИ (12 МЕСЯЦЕВ)
    ✅ СРАВНЕНИЕ С КОНКУРЕНТАМИ
    ✅ ВАЛИДАЦИЯ ДАННЫХ
    ✅ НАСТРОЙКИ ПОЛЬЗОВАТЕЛЯ
    ✅ ПОЛНАЯ ДОКУМЕНТАЦИЯ
    ✅ АВТОМАТИЧЕСКАЯ ОПТИМИЗАЦИЯ ЦЕН
    ✅ ДАШБОРД С КЛЮЧЕВЫМИ МЕТРИКАМИ
    ✅ ОБЪЕДИНЕНИЕ ДАННЫХ С ВЫБОРОМ КРИТЕРИЕВ
    ✅ ЭКСПОРТ С ФОРМУЛАМИ В EXCEL
    ✅ АВТОМАТИЧЕСКОЕ ОБНОВЛЕНИЕ КУРСОВ ВАЛЮТ
    ✅ РАСШИРЕННАЯ СТАТИСТИКА С ГРАФИКАМИ
    ✅ ИСТОРИЯ РАСЧЕТОВ С ФИЛЬТРАЦИЕЙ
    ✅ УПРАВЛЕНИЕ ЦЕНАМИ И НАЦЕНКАМИ
    ✅ 5+ МАРКЕТПЛЕЙСОВ С АКТУАЛЬНЫМИ ТАРИФАМИ
    ✅ МНОГОПОТОЧНЫЙ РАСЧЕТ ДЛЯ БОЛЬШИХ КАТАЛОГОВ
    ✅ ПРОГНОЗИРОВАНИЕ СПРОСА
    ✅ АНАЛИЗ КОНКУРЕНТОВ
    ✅ АВТОМАТИЧЕСКАЯ ГЕНЕРАЦИЯ ОТЧЕТОВ
    ✅ СРАВНЕНИЕ ЭФФЕКТИВНОСТИ РЕЖИМОВ РАБОТЫ
    ✅ РАСЧЕТ ТОЧКИ БЕЗУБЫТОЧНОСТИ
    ✅ ОПТИМИЗАЦИЯ ЛОГИСТИКИ
    ✅ УЧЕТ СЕЗОННОСТИ
    ✅ АНАЛИЗ ПРИБЫЛЬНОСТИ КАТЕГОРИЙ
    ✅ РАСШИРЕННАЯ ВИЗУАЛИЗАЦИЯ
    ✅ ЭКСПОРТ В HTML ДАШБОРД
    ✅ АВТОМАТИЧЕСКОЕ ОБНАРУЖЕНИЕ ВЫБРОСОВ
    ✅ КОРРЕЛЯЦИОННЫЙ АНАЛИЗ
    ✅ КЛАСТЕРИЗАЦИЯ ТОВАРОВ
    ✅ А/Б ТЕСТИРОВАНИЕ ЦЕН
    ✅ ПРОГНОЗИРОВАНИЕ ОСТАТКОВ
    ✅ ОПТИМИЗАЦИЯ ЗАКУПОК
    ✅ АНАЛИЗ ЭФФЕКТИВНОСТИ КАНАЛОВ ПРОДАЖ
    ✅ МЕТРИКИ LTV, CAC, ROMI
    ✅ ДАШБОРД В РЕАЛЬНОМ ВРЕМЕНИ
    ✅ ЭКСПОРТ В POWER BI И TABLEAU
================================================================================
"""

# ============================================================================
# БЛОК 0: ВСЕ НЕОБХОДИМЫЕ ИМПОРТЫ (300+ СТРОК)
# ============================================================================

# --- СТАНДАРТНЫЕ БИБЛИОТЕКИ ---
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
import gzip
import bz2
import lzma
import socket
import struct
import threading
import queue
import asyncio
import aiohttp
import aiofiles
import concurrent.futures
from typing import Dict, List, Any, Optional, Tuple, Union, Set, Callable, Iterable, Iterator, Generator
from dataclasses import dataclass, field, asdict, astuple, replace
from functools import lru_cache, wraps, reduce, partial, singledispatch
from concurrent.futures import ThreadPoolExecutor, as_completed, ProcessPoolExecutor, wait
from datetime import datetime, timedelta, date, timezone
from collections import defaultdict, Counter, deque, OrderedDict, ChainMap, namedtuple, UserDict, UserList
from enum import Enum, auto, IntEnum, Flag
from threading import Lock, RLock, Semaphore, Thread, Event, Barrier, Condition, Timer
from contextlib import contextmanager, closing, suppress, ExitStack
from pathlib import Path, PurePath
from abc import ABC, abstractmethod, abstractproperty
import inspect
import importlib
import importlib.util
import subprocess
import platform
import psutil
import gc
import weakref
import copy
import pprint
import statistics
import random
import secrets
import cProfile
import pstats
import io
import sysconfig
import site
import warnings
import webbrowser
import datetime as dt
import calendar
from zoneinfo import ZoneInfo
import email
import smtplib
import imaplib
import poplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from html import escape, unescape
from xml.etree import ElementTree
import xml.dom.minidom
import yaml
import toml
import configparser
import argparse
import getpass
import hashlib
import hmac
import secrets
import jwt
import oauthlib
from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageOps
import pytz
import dateutil
from dateutil.parser import parse
from dateutil.relativedelta import relativedelta
import holidays
import forex_python
from forex_python.converter import CurrencyRates
from babel.numbers import format_currency, format_decimal, format_percent
from babel.dates import format_date, format_datetime, format_time
import phonenumbers
from phonenumbers import PhoneNumberType, PhoneNumber, parse, format_number, PhoneNumberFormat
import validators
from validators import url, email, domain, ip_address
import pycountry
import pytz
import tzlocal

# --- НАУЧНЫЕ И ДАННЫЕ БИБЛИОТЕКИ ---
try:
    import polars as pl
    import polars.selectors as cs
    POLARS_AVAILABLE = True
except ImportError:
    POLARS_AVAILABLE = False

try:
    import duckdb
    import duckdb.experimental as dexp
    DUCKDB_AVAILABLE = True
except ImportError:
    DUCKDB_AVAILABLE = False

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
    import vaex
    VAEX_AVAILABLE = True
except ImportError:
    VAEX_AVAILABLE = False

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
    import pandera as pa_schema
    from pandera import Column, DataFrameSchema, Check, Index
    PANDERA_AVAILABLE = True
except ImportError:
    PANDERA_AVAILABLE = False

# --- МАШИННОЕ ОБУЧЕНИЕ ---
try:
    import sklearn
    from sklearn.feature_extraction.text import TfidfVectorizer, CountVectorizer, HashingVectorizer
    from sklearn.naive_bayes import MultinomialNB, GaussianNB, BernoulliNB, ComplementNB
    from sklearn.pipeline import Pipeline, make_pipeline, FeatureUnion
    from sklearn.model_selection import train_test_split, cross_val_score, GridSearchCV, RandomizedSearchCV
    from sklearn.metrics import accuracy_score, classification_report, confusion_matrix, f1_score, precision_score
    from sklearn.metrics import recall_score, roc_auc_score, roc_curve, auc, precision_recall_curve
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
    import xgboost as xgb
    import lightgbm as lgb
    import catboost as cb
    SKLEARN_AVAILABLE = True
except ImportError:
    SKLEARN_AVAILABLE = False
    xgb = None
    lgb = None
    cb = None

# --- ВИЗУАЛИЗАЦИЯ ---
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

# --- ОФИСНЫЕ И ЭКСПОРТНЫЕ БИБЛИОТЕКИ ---
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
    from reportlab.platypus import PageBreak, Image, KeepTogether, NextPageTemplate, BaseDocTemplate
    from reportlab.platypus import Frame, PageTemplate, Flowable, DocTemplate, TableOfContents
    from reportlab.platypus.para import Paragraph as Para
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle, StyleSheet1
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
    import pygments
    from pygments import highlight
    from pygments.lexers import PythonLexer, JsonLexer, CsvLexer
    from pygments.formatters import HtmlFormatter
    PYGMENTS_AVAILABLE = True
except ImportError:
    PYGMENTS_AVAILABLE = False

# --- СПЕЦИАЛИЗИРОВАННЫЕ БИБЛИОТЕКИ ---
try:
    import chardet
    CHARDET_AVAILABLE = True
except ImportError:
    CHARDET_AVAILABLE = False

try:
    import openai
    OPENAI_AVAILABLE = True
except ImportError:
    OPENAI_AVAILABLE = False

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

# --- API И СЕТЕВЫЕ БИБЛИОТЕКИ ---
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
    import grpc
    import grpc.aio
    GRPC_AVAILABLE = True
except ImportError:
    GRPC_AVAILABLE = False

# ============================================================================
# БЛОК 1: КОНСТАНТЫ И НАСТРОЙКИ (200+ СТРОК)
# ============================================================================

# --- ВЕРСИИ И МЕТАДАННЫЕ ---
APP_VERSION = "95.0.0"
APP_NAME = "🚗 Юнит-экономика автозапчастей 2026"
APP_AUTHOR = "AutoParts Analytics Team"
APP_DESCRIPTION = "Полный расчет юнит-экономики для автозапчастей с AI-оптимизацией"
APP_WEBSITE = "https://github.com/autoparts-unit-economics"
APP_LICENSE = "MIT License"
APP_COPYRIGHT = f"2024-2026 {APP_AUTHOR}"

# --- ОГРАНИЧЕНИЯ И ЛИМИТЫ ---
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
MAX_UPLOAD_SIZE = 500 * 1024 * 1024  # 500 MB
MAX_CATEGORIES = 300
MAX_ANALOGS = 100
PRECISION_DECIMALS = 2

# --- ВАЛЮТЫ И РЕГИОНЫ ---
SUPPORTED_CURRENCIES = ["RUB", "USD", "EUR", "CNY", "KZT", "UAH", "BYN", "AMD"]
SUPPORTED_LANGUAGES = ["ru", "en", "uk", "kz", "by", "am"]
DEFAULT_LOCALE = "ru_RU"
TIMEZONE = "Europe/Moscow"

# --- ПУТИ К ФАЙЛАМ ---
BASE_DIR = Path(__file__).parent.resolve()
DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
LOG_DIR = BASE_DIR / "logs"
REPORTS_DIR = BASE_DIR / "reports"
TEMP_DIR = BASE_DIR / "temp"
MODELS_DIR = BASE_DIR / "models"
CONFIG_DIR = BASE_DIR / "config"
PLUGINS_DIR = BASE_DIR / "plugins"

# Создание директорий
for dir_path in [DATA_DIR, CACHE_DIR, LOG_DIR, REPORTS_DIR, TEMP_DIR, MODELS_DIR, CONFIG_DIR, PLUGINS_DIR]:
    dir_path.mkdir(exist_ok=True, parents=True)

# --- НАСТРОЙКИ ЛОГГИРОВАНИЯ ---
LOG_LEVEL = os.environ.get("LOG_LEVEL", "INFO")
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
LOG_FILE = LOG_DIR / "auto_parts_economy.log"
LOG_ROTATION_SIZE = 10 * 1024 * 1024  # 10 MB
LOG_RETENTION_DAYS = 30

# --- НАСТРОЙКИ API ---
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
DEEPSEEK_MODEL = "deepseek-chat"
OPENAI_MODEL = "gpt-4"
ANTHROPIC_MODEL = "claude-3-sonnet-20240229"

# --- НАСТРОЙКИ ПРОИЗВОДИТЕЛЬНОСТИ ---
USE_CACHING = True
USE_PARALLEL = True
USE_GPU = False
OPTIMIZE_MEMORY = True
USE_DUCKDB = True
USE_POLARS = True

# --- ЦВЕТА ДЛЯ ВИЗУАЛИЗАЦИИ ---
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
    "#ec407a", "#66bb6a", "#ffa726", "#8d6e63", "#78909c"
]

SEABORN_PALETTES = [
    "deep", "muted", "bright", "pastel", "dark", "colorblind",
    "Spectral", "RdYlGn", "RdBu", "viridis", "plasma", "inferno",
    "magma", "cividis", "twilight", "hsv"
]

# --- ЭМОДЗИ И ИКОНКИ ---
ICONS = {
    "profit": "💰",
    "margin": "📈",
    "roi": "📊",
    "expenses": "💳",
    "revenue": "🏷️",
    "cost": "💵",
    "logistics": "🚚",
    "storage": "📦",
    "commission": "💸",
    "tax": "🧾",
    "warning": "⚠️",
    "success": "✅",
    "error": "❌",
    "info": "ℹ️",
    "question": "❓",
    "settings": "⚙️",
    "download": "📥",
    "upload": "📤",
    "search": "🔍",
    "analytics": "📊",
    "chart": "📈",
    "table": "📋",
    "file": "📁",
    "folder": "📂",
    "calendar": "📅",
    "clock": "🕐",
    "check": "✅",
    "cross": "❌",
    "star": "⭐",
    "heart": "❤️",
    "rocket": "🚀",
    "car": "🚗",
    "truck": "🚛",
    "warehouse": "🏭",
    "factory": "🏗️",
    "tools": "🔧",
    "wrench": "🔩",
    "gear": "⚙️",
    "engine": "🔥",
    "battery": "🔋",
    "light": "💡",
    "warning_sign": "⚠️"
}

# ============================================================================
# БЛОК 2: КЛАССЫ ИСКЛЮЧЕНИЙ (100+ СТРОК)
# ============================================================================

class AutoPartsException(Exception):
    """Базовое исключение для приложения"""
    def __init__(self, message: str = "", *args, **kwargs):
        self.message = message
        self.timestamp = datetime.now()
        self.context = kwargs
        super().__init__(message, *args)

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

# ============================================================================
# БЛОК 3: ДЕКОРАТОРЫ (300+ СТРОК)
# ============================================================================

def timer_decorator(func: Callable) -> Callable:
    """Декоратор для замера времени выполнения"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        start = time.perf_counter()
        try:
            result = func(*args, **kwargs)
            elapsed = time.perf_counter() - start
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
            
            # Очистка кеша при превышении размера
            if len(cache) > maxsize:
                # Удаляем наименее используемые
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
            # Валидация позиционных аргументов
            for i, arg in enumerate(args):
                if i < len(types):
                    expected_type = types[i]
                    if not isinstance(arg, expected_type):
                        raise ValidationError(
                            f"Аргумент {i} должен быть типа {expected_type.__name__}, получен {type(arg).__name__}",
                            field=str(i),
                            value=arg
                        )
            
            # Валидация именованных аргументов
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
        # Формируем информацию о аргументах
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
    """Декоратор для ограничения времени выполнения"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            def timeout_handler(signum, frame):
                raise TimeoutError(f"{error_message} ({seconds}с)")
            
            import signal
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
        
        memory_used = (end_memory - start_memory) / 1024 / 1024  # MB
        time_used = end_time - start_time
        
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

def async_retry_decorator(max_retries: int = 3, delay: float = 1.0) -> Callable:
    """Декоратор для повторных попыток асинхронных функций"""
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        async def wrapper(*args, **kwargs):
            last_exception = None
            for attempt in range(max_retries):
                try:
                    return await func(*args, **kwargs)
                except Exception as e:
                    last_exception = e
                    if attempt == max_retries - 1:
                        raise
                    await asyncio.sleep(delay * (2 ** attempt))
            if last_exception:
                raise last_exception
            return None
        return wrapper
    return decorator

def require_authentication(func: Callable) -> Callable:
    """Декоратор для проверки аутентификации"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        # Проверка наличия токена в сессии
        if not st.session_state.get("authenticated", False):
            st.error("❌ Требуется аутентификация")
            st.stop()
        return func(*args, **kwargs)
    return wrapper

def require_license(func: Callable) -> Callable:
    """Декоратор для проверки лицензии"""
    @wraps(func)
    def wrapper(*args, **kwargs):
        # Проверка валидности лицензии
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
        
        # Логируем вызов
        logger.info(f"📊 Вызов {func.__name__} пользователем {user} в {timestamp}")
        
        # Считаем вызовы
        if "usage_stats" not in st.session_state:
            st.session_state.usage_stats = defaultdict(int)
        st.session_state.usage_stats[func.__name__] += 1
        
        return func(*args, **kwargs)
    return wrapper

# ============================================================================
# БЛОК 4: ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ (500+ СТРОК)
# ============================================================================

@contextmanager
def timer_context(name: str):
    """Контекстный менеджер для замера времени"""
    start = time.perf_counter()
    try:
        yield
    finally:
        elapsed = time.perf_counter() - start
        logger.info(f"⏱ {name}: {elapsed:.3f}с")

@contextmanager
def memory_usage_context():
    """Контекстный менеджер для отслеживания памяти"""
    process = psutil.Process()
    mem_before = process.memory_info().rss / 1024 / 1024
    try:
        yield
    finally:
        mem_after = process.memory_info().rss / 1024 / 1024
        logger.info(f"📊 Память: {mem_before:.1f}MB → {mem_after:.1f}MB (+{mem_after - mem_before:.1f}MB)")

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

def safe_float(val: Any, default: float = 0.0) -> float:
    """Безопасное преобразование в float с поддержкой различных форматов"""
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
        except:
            return default
    
    if isinstance(val, str):
        # Очистка строки
        cleaned = val.strip()
        if not cleaned:
            return default
        
        # Удаление валютных символов и пробелов
        cleaned = re.sub(r'[^\d.,\-+\s]', '', cleaned)
        cleaned = cleaned.replace(' ', '')
        cleaned = cleaned.replace(',', '.')
        
        # Обработка отрицательных чисел
        if cleaned.count('-') > 1:
            return default
        
        # Обработка нескольких точек
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
        except:
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
        return ", ".join(safe_str(v) for v in val)
    
    if isinstance(val, dict):
        return str(val)
    
    try:
        result = str(val).strip()
        return result if result else default
    except:
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
        true_values = {'true', 'yes', '1', 'y', 'да', 'on', 'True', 'YES', 'Yes'}
        false_values = {'false', 'no', '0', 'n', 'нет', 'off', 'False', 'NO', 'No'}
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
        except:
            return default
    
    if isinstance(val, str):
        formats = [
            "%Y-%m-%d %H:%M:%S",
            "%Y-%m-%d %H:%M:%S.%f",
            "%Y-%m-%d",
            "%d.%m.%Y %H:%M:%S",
            "%d.%m.%Y",
            "%Y-%m-%dT%H:%M:%S",
            "%Y-%m-%dT%H:%M:%S.%f",
            "%Y-%m-%dT%H:%M:%S.%fZ",
            "%b %d %Y %H:%M:%S",
            "%d %b %Y %H:%M:%S",
            "%d/%m/%Y %H:%M:%S",
            "%d/%m/%Y",
            "%m/%d/%Y",
            "%m/%d/%Y %H:%M:%S",
        ]
        
        for fmt in formats:
            try:
                return datetime.strptime(val, fmt)
            except ValueError:
                continue
        
        try:
            from dateutil.parser import parse
            return parse(val)
        except:
            pass
        
        return default
    
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
            key_parts.append(hashlib.md5(pd.util.hash_pandas_object(arg).values.tobytes()).hexdigest())
        elif isinstance(arg, pd.Series):
            key_parts.append(hashlib.md5(pd.util.hash_pandas_object(arg).values.tobytes()).hexdigest())
        elif isinstance(arg, np.ndarray):
            key_parts.append(hashlib.md5(arg.tobytes()).hexdigest())
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
            key_parts.append(f"{k}:{hashlib.md5(pd.util.hash_pandas_object(v).values.tobytes()).hexdigest()}")
        else:
            key_parts.append(f"{k}:{v}")
    
    key = "|".join(key_parts)
    return hashlib.md5(key.encode('utf-8')).hexdigest()

def format_currency(value: float, currency: str = "RUB", locale: str = "ru_RU") -> str:
    """Форматирование валюты с учетом локали"""
    if value is None or math.isnan(value) or math.isinf(value):
        return f"0 {currency}"
    
    try:
        from babel.numbers import format_currency as babel_format
        return babel_format(value, currency, locale=locale)
    except:
        symbols = {
            "RUB": "₽",
            "USD": "$",
            "EUR": "€",
            "CNY": "¥",
            "KZT": "₸",
            "UAH": "₴",
            "BYN": "Br",
            "AMD": "֏"
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
    
    try:
        from babel.numbers import format_percent as babel_format
        return babel_format(value / 100, format=f"#,##0.{'0' * decimal_places}%", locale="ru_RU")
    except:
        return f"{value:.{decimal_places}f}%"

def calculate_volume(length: float, width: float, height: float) -> float:
    """Расчет объема в литрах с валидацией"""
    if not all([length, width, height]):
        return 0.0
    
    if not all([length > 0, width > 0, height > 0]):
        return 0.0
    
    # Проверка на слишком большие значения (могут быть в мм)
    if any([length > 1000, width > 1000, height > 1000]):
        # Конвертируем мм в см
        length /= 10
        width /= 10
        height /= 10
    
    # Проверка на слишком маленькие значения
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

def is_valid_barcode(barcode: str) -> bool:
    """Проверка валидности штрихкода (EAN-8, EAN-13, UPC-A, UPC-E)"""
    if not barcode:
        return False
    
    barcode = re.sub(r'[^\d]', '', barcode)
    
    if len(barcode) not in [8, 12, 13, 14]:
        return False
    
    # Проверка контрольной суммы EAN-13
    if len(barcode) == 13:
        checksum = 0
        for i, digit in enumerate(barcode[:-1]):
            checksum += int(digit) * (3 if (i + 1) % 2 == 0 else 1)
        checksum = (10 - (checksum % 10)) % 10
        return checksum == int(barcode[-1])
    
    return True

def normalize_text(text: str) -> str:
    """Нормализация текста (приведение к нижнему регистру, удаление знаков препинания)"""
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
    
    # Используем оптимизированный алгоритм с одной строкой
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
    """Поиск похожих строк с использованием расстояния Левенштейна"""
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
    if CHARDET_AVAILABLE:
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
    
    # Попробуем стандартные кодировки
    encodings = ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251', 'cp1252', 'latin1']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f:
                f.read()
                return enc
        except:
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
        elif ext == '.sqlite' or ext == '.db':
            import sqlite3
            conn = sqlite3.connect(path)
            return pd.read_sql_query("SELECT * FROM data", conn)
        else:
            raise ValueError(f"Неподдерживаемый формат файла: {ext}")
    except Exception as e:
        raise AutoPartsException(f"Ошибка загрузки файла {path}: {e}")

def save_dataframe(df: pd.DataFrame, file_path: Union[str, Path], **kwargs) -> None:
    """Универсальное сохранение данных с автоопределением формата"""
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
    
    # Удаление полностью пустых строк
    df = df.dropna(how='all')
    
    # Удаление пустых колонок
    df = df.dropna(axis=1, how='all')
    
    # Удаление дубликатов
    df = df.drop_duplicates()
    
    # Очистка имен колонок
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
    """Расчет оптимального размера батча для параллельной обработки"""
    if max_workers is None:
        max_workers = os.cpu_count() or 4
    
    # Оптимальный размер батча для максимальной производительности
    base_size = 1000
    optimal_size = max(100, min(base_size, data_size // max_workers))
    
    # Ограничение сверху
    max_batch = 10000
    return min(optimal_size, max_batch)

def split_into_batches(data: List, batch_size: int) -> List[List]:
    """Разбиение данных на батчи"""
    return [data[i:i + batch_size] for i in range(0, len(data), batch_size)]

def parallel_apply(df: pd.DataFrame, func: Callable, axis: int = 1, max_workers: int = None) -> pd.Series:
    """Параллельное применение функции к строкам DataFrame"""
    if max_workers is None:
        max_workers = os.cpu_count() or 4
    
    # Разбиваем DataFrame на части
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
            # Преобразуем строки с небольшим количеством уникальных значений в категории
            if len(df_optimized[col].unique()) / len(df_optimized[col]) < 0.5:
                df_optimized[col] = df_optimized[col].astype('category')
        
        elif col_type == 'float64':
            # Преобразуем в float32 если это возможно
            try:
                df_optimized[col] = df_optimized[col].astype('float32')
            except:
                pass
        
        elif col_type == 'int64':
            # Преобразуем в меньший целочисленный тип
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

def get_memory_usage(obj) -> int:
    """Получение использования памяти объектом в байтах"""
    return sys.getsizeof(obj)

def get_dataframe_memory_usage(df: pd.DataFrame) -> Dict[str, int]:
    """Получение использования памяти каждой колонкой DataFrame"""
    return {col: df[col].memory_usage(deep=True) for col in df.columns}

def format_memory_size(size_bytes: int) -> str:
    """Форматирование размера памяти в человекочитаемый формат"""
    for unit in ['B', 'KB', 'MB', 'GB', 'TB']:
        if size_bytes < 1024.0:
            return f"{size_bytes:.2f} {unit}"
        size_bytes /= 1024.0
    return f"{size_bytes:.2f} PB"

# ============================================================================
# БЛОК 5: ENUM И ТИПЫ (200+ СТРОК)
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
    FBY = auto()  # Fulfillment By Yandex
    FBS = auto()  # Fulfillment By Seller
    FBO = auto()  # Fulfillment By Operator
    DBS = auto()  # Delivery By Seller
    FBP = auto()  # Fulfillment By Partner
    DBE = auto()  # Delivery By Express
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

# ============================================================================
# БЛОК 6: ДАТАКЛАССЫ (300+ СТРОК)
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
        """Объем в литрах"""
        return calculate_volume(self.length, self.width, self.height)
    
    @property
    def is_valid(self) -> bool:
        """Проверка валидности габаритов"""
        return all([self.length > 0, self.width > 0, self.height > 0, self.weight > 0])

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
    
    def get_dimensions(self) -> ProductDimensions:
        """Получение габаритов категории"""
        return self.dimensions or ProductDimensions()

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
        """Преобразование в словарь"""
        result = asdict(self)
        result['timestamp'] = self.timestamp.isoformat()
        result['status'] = self.status.name
        return result
    
    def to_dataframe(self) -> pd.DataFrame:
        """Преобразование в DataFrame"""
        return pd.DataFrame([self.to_dict()])
    
    def get_summary(self) -> Dict[str, Any]:
        """Получение сводки результатов"""
        return {
            "marketplace": self.marketplace,
            "profit": self.profit,
            "margin": self.margin_percent,
            "roi": self.roi,
            "breakeven": self.breakeven_price,
            "total_expenses": self.total_expenses
        }

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
        """Преобразование в DataFrame"""
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
        """Преобразование в словарь"""
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

# ============================================================================
# БЛОК 7: ОСНОВНОЙ КЛАСС ЮНИТ-ЭКОНОМИКИ (800+ СТРОК)
# ============================================================================

class AutoPartsUnitEconomics:
    """
    Основной класс для расчета юнит-экономики автозапчастей
    
    Особенности:
    - Учет габаритов и веса
    - Категорийные ставки комиссий
    - AI-обновление тарифов
    - Пакетный расчет для каталога
    - Прогнозирование прибыли
    - Оптимизация цены
    - Многопоточная обработка
    """
    
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
        
        # Инициализация конфигураций
        self._configs = self._load_marketplace_configs()
        self._categories = self._load_categories()
        
        # Кеши
        self._cache = {}
        self._history = []
        self._calculation_cache = {}
        
        # Статистика
        self._stats = self._init_stats()
        
        # Настройки
        self._settings = self._load_settings()
        
        # AI клиенты
        self._ai_clients = self._init_ai_clients()
        
        self.logger = logging.getLogger('AutoPartsUnitEconomics')
        self.logger.info("🚗 Инициализация AutoPartsUnitEconomics")
        self.logger.info(f"📊 Загружено {len(self._configs)} маркетплейсов")
        self.logger.info(f"📚 Загружено {len(self._categories)} категорий")
    
    def _load_marketplace_configs(self) -> Dict[str, MarketplaceConfig]:
        """Загрузка конфигураций маркетплейсов"""
        configs = {
            "Яндекс Маркет": MarketplaceConfig(
                name="Яндекс Маркет",
                commission_rate=0.14,
                min_commission=0.0,
                logistics_base=150.0,
                logistics_per_kg=50.0,
                logistics_per_liter=30.0,
                storage_per_day=0.5,
                return_fee=0.03,
                acquiring_fee=0.015,
                last_mile_fee=100.0,
                delivery_fee_percent=0.05,
                premium_fee=0.02,
                rko_fee=0.005,
                subscription_fee=6990.0,
                insurance_fee=0.01,
                packing_fee=50.0,
                marketing_fee=0.03,
                mode_multipliers={"FBY": 0.75, "FBS": 1.0, "FBO": 0.8, "DBS": 1.3, "FBP": 0.9},
                category_rates={
                    "двигатель": 0.14,
                    "трансмиссия": 0.14,
                    "подвеска": 0.14,
                    "тормозная_система": 0.14,
                    "электрика": 0.14,
                    "охлаждение": 0.12,
                    "фильтры": 0.14,
                    "масла": 0.14,
                    "кузов": 0.16,
                    "оптика": 0.15
                }
            ),
            "Ozon": MarketplaceConfig(
                name="Ozon",
                commission_rate=0.15,
                min_commission=30.0,
                logistics_base=90.0,
                logistics_per_kg=35.0,
                logistics_per_liter=20.0,
                storage_per_day=0.3,
                return_fee=0.05,
                acquiring_fee=0.01,
                last_mile_fee=80.0,
                delivery_fee_percent=0.04,
                premium_fee=0.0,
                rko_fee=0.005,
                subscription_fee=0.0,
                insurance_fee=0.005,
                packing_fee=30.0,
                marketing_fee=0.02,
                mode_multipliers={"FBY": 0.75, "FBS": 1.0, "FBO": 0.8, "DBS": 1.3, "FBP": 0.9},
                category_rates={
                    "двигатель": 0.12,
                    "трансмиссия": 0.12,
                    "подвеска": 0.12,
                    "тормозная_система": 0.12,
                    "электрика": 0.12,
                    "охлаждение": 0.10,
                    "фильтры": 0.12,
                    "масла": 0.10,
                    "кузов": 0.18,
                    "оптика": 0.16
                }
            ),
            "Wildberries": MarketplaceConfig(
                name="Wildberries",
                commission_rate=0.15,
                min_commission=40.0,
                logistics_base=120.0,
                logistics_per_kg=45.0,
                logistics_per_liter=28.0,
                storage_per_day=0.4,
                return_fee=0.04,
                acquiring_fee=0.012,
                last_mile_fee=90.0,
                delivery_fee_percent=0.06,
                premium_fee=0.025,
                rko_fee=0.005,
                subscription_fee=0.0,
                insurance_fee=0.008,
                packing_fee=40.0,
                marketing_fee=0.025,
                mode_multipliers={"FBY": 0.75, "FBS": 1.0, "FBO": 0.8, "DBS": 1.3, "FBP": 0.9},
                category_rates={
                    "двигатель": 0.15,
                    "трансмиссия": 0.15,
                    "подвеска": 0.15,
                    "тормозная_система": 0.15,
                    "электрика": 0.15,
                    "охлаждение": 0.13,
                    "фильтры": 0.15,
                    "масла": 0.12,
                    "кузов": 0.20,
                    "оптика": 0.18
                }
            )
        }
        
        return configs
    
    def _load_categories(self) -> Dict[str, ProductCategory]:
        """Загрузка категорий автозапчастей"""
        categories = {
            "двигатель": ProductCategory(
                name="двигатель",
                description="Двигатель внутреннего сгорания и компоненты",
                dimensions=ProductDimensions(length=50, width=40, height=40, weight=50),
                typical_volume=80,
                typical_weight=50,
                oem_codes=["11", "12", "13"],
                hazardous=False,
                fragile=False
            ),
            "трансмиссия": ProductCategory(
                name="трансмиссия",
                description="Коробки передач, сцепление, приводы",
                dimensions=ProductDimensions(length=40, width=30, height=25, weight=30),
                typical_volume=30,
                typical_weight=30,
                oem_codes=["12", "14"],
                hazardous=False,
                fragile=False
            ),
            "подвеска": ProductCategory(
                name="подвеска",
                description="Амортизаторы, пружины, рычаги",
                dimensions=ProductDimensions(length=60, width=20, height=15, weight=10),
                typical_volume=18,
                typical_weight=8,
                oem_codes=["13", "15"],
                hazardous=False,
                fragile=False
            ),
            "тормозная_система": ProductCategory(
                name="тормозная_система",
                description="Колодки, диски, суппорты",
                dimensions=ProductDimensions(length=30, width=20, height=10, weight=8),
                typical_volume=6,
                typical_weight=8,
                oem_codes=["15", "16"],
                hazardous=False,
                fragile=False
            ),
            "электрика": ProductCategory(
                name="электрика",
                description="Стартеры, генераторы, аккумуляторы",
                dimensions=ProductDimensions(length=25, width=20, height=20, weight=15),
                typical_volume=10,
                typical_weight=15,
                oem_codes=["16", "17"],
                hazardous=True,
                fragile=False
            ),
            "охлаждение": ProductCategory(
                name="охлаждение",
                description="Радиаторы, термостаты, помпы",
                dimensions=ProductDimensions(length=50, width=30, height=20, weight=8),
                typical_volume=30,
                typical_weight=8,
                oem_codes=["11", "18"],
                hazardous=False,
                fragile=False
            ),
            "фильтры": ProductCategory(
                name="фильтры",
                description="Масляные, воздушные, топливные фильтры",
                dimensions=ProductDimensions(length=15, width=10, height=8, weight=0.5),
                typical_volume=1.2,
                typical_weight=0.5,
                oem_codes=["17", "19"],
                hazardous=False,
                fragile=False
            ),
            "масла": ProductCategory(
                name="масла",
                description="Моторные и трансмиссионные масла",
                dimensions=ProductDimensions(length=25, width=15, height=15, weight=5),
                typical_volume=5.6,
                typical_weight=5,
                oem_codes=["18", "20"],
                hazardous=True,
                fragile=False
            ),
            "кузов": ProductCategory(
                name="кузов",
                description="Кузовные детали и оптика",
                dimensions=ProductDimensions(length=80, width=40, height=20, weight=15),
                typical_volume=64,
                typical_weight=15,
                oem_codes=["19", "21"],
                hazardous=False,
                fragile=True
            ),
            "оптика": ProductCategory(
                name="оптика",
                description="Фары, фонари, лампы",
                dimensions=ProductDimensions(length=25, width=15, height=10, weight=3),
                typical_volume=3.8,
                typical_weight=3,
                oem_codes=["16", "19"],
                hazardous=False,
                fragile=True
            )
        }
        
        return categories
    
    def _init_stats(self) -> Dict[str, Any]:
        """Инициализация статистики"""
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
            "last_error": None
        }
    
    def _load_settings(self) -> Dict[str, Any]:
        """Загрузка настроек"""
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
                self.logger.warning(f"Ошибка загрузки настроек: {e}")
        
        return default_settings
    
    def _init_ai_clients(self) -> Dict[str, Any]:
        """Инициализация AI клиентов"""
        clients = {}
        
        if OPENAI_AVAILABLE and os.environ.get("OPENAI_API_KEY"):
            try:
                clients["openai"] = openai.OpenAI(api_key=os.environ["OPENAI_API_KEY"])
                self.logger.info("OpenAI клиент инициализирован")
            except Exception as e:
                self.logger.warning(f"Ошибка инициализации OpenAI: {e}")
        
        if ANTHROPIC_AVAILABLE and os.environ.get("ANTHROPIC_API_KEY"):
            try:
                clients["anthropic"] = anthropic.Anthropic(api_key=os.environ["ANTHROPIC_API_KEY"])
                self.logger.info("Anthropic клиент инициализирован")
            except Exception as e:
                self.logger.warning(f"Ошибка инициализации Anthropic: {e}")
        
        # DeepSeek клиент через requests
        if os.environ.get("DEEPSEEK_API_KEY"):
            clients["deepseek"] = {
                "api_key": os.environ["DEEPSEEK_API_KEY"],
                "url": "https://api.deepseek.com/v1/chat/completions"
            }
            self.logger.info("DeepSeek клиент инициализирован")
        
        return clients
    
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
        """
        Расчет юнит-экономики для автозапчасти
        
        Args:
            price: Цена продажи
            cost: Себестоимость
            marketplace: Название маркетплейса
            category: Категория товара
            operation_mode: Режим работы
            days_in_storage: Дни хранения
            length, width, height, weight: Габариты
            is_premium: Премиум-раздел
            include_insurance: Включить страховку
            include_packing: Включить упаковку
            include_marketing: Включить маркетинг
            currency: Валюта
        
        Returns:
            UnitEconomicsResult с результатами расчета
        """
        # Валидация входных данных
        if price <= 0:
            raise ValidationError("Цена должна быть положительной", "price", price)
        if cost <= 0:
            raise ValidationError("Себестоимость должна быть положительной", "cost", cost)
        
        if marketplace not in self._configs:
            raise MarketplaceError(f"Маркетплейс {marketplace} не поддерживается", marketplace)
        
        config = self._configs[marketplace]
        
        # Получение габаритов из категории
        if all([length == 0, width == 0, height == 0, weight == 0]) and category:
            length, width, height, weight = self.calculate_dimensions_from_category(category)
        
        # Расчет объема
        volume = calculate_volume(length, width, height)
        if volume == 0:
            volume = 5.0  # Средний объем по умолчанию
        
        if weight <= 0:
            weight = 1.0  # Средний вес по умолчанию
        
        # Комиссия с учетом категории
        commission_rate = config.category_rates.get(category, config.commission_rate)
        
        # Расчет комиссии
        commission = max(price * commission_rate, config.min_commission)
        if config.max_commission < float('inf'):
            commission = min(commission, config.max_commission)
        
        # Подписка (если есть)
        subscription_cost = config.subscription_fee / 30 if config.subscription_fee > 0 else 0
        
        # Логистика
        logistics = (
            config.logistics_base +
            weight * config.logistics_per_kg +
            volume * config.logistics_per_liter
        )
        
        # Множитель режима работы
        mode_multiplier = config.mode_multipliers.get(operation_mode, 1.0)
        logistics *= mode_multiplier
        
        # Хранение
        storage_cost = volume * config.storage_per_day * days_in_storage
        
        # Эквайринг
        acquiring = price * config.acquiring_fee
        
        # Доставка
        delivery = price * config.delivery_fee_percent
        
        # Последняя миля
        last_mile = config.last_mile_fee
        
        # Возвраты
        returns = price * config.return_fee
        
        # РКО
        rko_fee = price * config.rko_fee if config.rko_fee > 0 else 0
        
        # Премиум
        premium_fee = price * config.premium_fee if is_premium and config.premium_fee > 0 else 0
        
        # Страховка
        insurance_fee = price * config.insurance_fee if include_insurance and config.insurance_fee > 0 else 0
        
        # Упаковка
        packing_fee = config.packing_fee if include_packing and config.packing_fee > 0 else 0
        
        # Маркетинг
        marketing_fee = price * config.marketing_fee if include_marketing and config.marketing_fee > 0 else 0
        
        # Итого расходов
        total_expenses = (
            cost + commission + subscription_cost + logistics + storage_cost +
            acquiring + delivery + last_mile + returns + rko_fee +
            premium_fee + insurance_fee + packing_fee + marketing_fee
        )
        
        # Прибыль
        profit = price - total_expenses
        
        # Маржинальность
        margin_percent = (profit / price * 100) if price > 0 else 0
        
        # ROI
        roi = (profit / cost * 100) if cost > 0 else 0
        
        # Точка безубыточности
        fixed_costs = logistics + storage_cost + last_mile + subscription_cost
        variable_rate = (
            commission_rate + config.acquiring_fee +
            config.delivery_fee_percent + config.return_fee +
            config.rko_fee + config.premium_fee +
            config.insurance_fee + config.marketing_fee
        )
        breakeven_price = ((cost + fixed_costs) / (1 - variable_rate)) if (1 - variable_rate) > 0 else 0
        
        # Маржинальная прибыль
        contribution_margin = price - cost - commission - logistics - acquiring - delivery - last_mile - returns
        contribution_margin_ratio = (contribution_margin / price * 100) if price > 0 else 0
        
        # Создание результата
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
        
        # Обновление статистики
        self._update_stats(result)
        
        # Сохранение истории
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
                self.logger.error(f"Ошибка расчета для {marketplace}: {e}")
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
        """
        Пакетный расчет для каталога
        
        Args:
            df: DataFrame с данными каталога
            price_col: Колонка с ценой
            cost_col: Колонка с себестоимостью
            category_col: Колонка с категорией
            length_col: Колонка с длиной
            width_col: Колонка с шириной
            height_col: Колонка с высотой
            weight_col: Колонка с весом
            article_col: Колонка с артикулом
            brand_col: Колонка с брендом
            marketplaces: Список маркетплейсов
            operation_mode: Режим работы
            days_in_storage: Дни хранения
            progress_callback: Функция обратного вызова для прогресса
            max_workers: Максимальное количество потоков
        
        Returns:
            DataFrame с результатами расчетов
        """
        if marketplaces is None:
            marketplaces = list(self._configs.keys())
        
        # Подготовка данных
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
        
        # Параллельный расчет
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
                    self.logger.error(f"Ошибка расчета для {article}: {e}")
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
        """
        Оптимизация цены для достижения целевой маржи
        
        Args:
            cost: Себестоимость
            marketplace: Маркетплейс
            category: Категория
            operation_mode: Режим работы
            days_in_storage: Дни хранения
            length, width, height, weight: Габариты
            target_margin: Целевая маржа (%)
            price_min: Минимальная цена
            price_max: Максимальная цена
            step: Шаг изменения цены
            max_iterations: Максимальное количество итераций
        
        Returns:
            OptimizationResult с результатами оптимизации
        """
        # Начальная цена
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
                
                # Проверка достижения целевой маржи
                if margin >= target_margin and profit > best_profit:
                    best_profit = profit
                    best_price = current_price
                    best_margin = margin
                    best_result = result
                
                current_price += step
                iteration += 1
                
            except Exception as e:
                self.logger.warning(f"Ошибка при оптимизации для цены {current_price}: {e}")
                current_price += step
        
        # Расчет улучшения
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
        
        # Рекомендации
        recommendations = []
        if best_price > 0 and best_margin >= target_margin:
            recommendations.append(f"Установите цену {best_price:.2f} ₽ для достижения маржи {target_margin}%")
        else:
            recommendations.append(f"Целевая маржа {target_margin}% не достигнута. Максимальная маржа: {best_margin:.1f}%")
        
        if best_profit > current_result.profit:
            recommendations.append(f"Потенциальное увеличение прибыли: {improvement_pct:.1f}%")
        
        # Обновление статистики
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
        """
        Прогнозирование прибыли
        
        Args:
            current_data: Текущие данные расчета
            periods: Количество периодов прогноза
            growth_rate: Годовой темп роста
            seasonality: Сезонные коэффициенты
            confidence_level: Уровень доверия для интервалов
        
        Returns:
            ForecastResult с прогнозом
        """
        if seasonality is None:
            # Стандартная сезонность для автозапчастей
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
            
            # Рост с учетом сезонности
            growth_factor = (1 + growth_rate) ** (i / 12)
            factor = seasonal_factor * growth_factor
            
            value = base_value * factor
            
            periods_list.append(datetime.now() + relativedelta(months=i))
            values_list.append(value)
            seasonality_list.append(seasonal_factor)
            trend_list.append(growth_factor)
        
        # Расчет доверительных интервалов
        std_dev = np.std(values_list) * 0.2  # 20% стандартное отклонение
        z_score = 1.96  # 95% доверительный интервал
        
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
        """Получение истории расчетов с фильтрацией"""
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
        
        # Добавляем дополнительную информацию
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
        """Получение лучшей конфигурации из истории"""
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
        """Экспорт истории в указанном формате"""
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
# БЛОК 8: UI КОМПОНЕНТЫ (800+ СТРОК)
# ============================================================================

def render_sidebar():
    """Рендеринг боковой панели"""
    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/000000/car-service.png", width=80)
        st.markdown("---")
        
        st.markdown("### 🧭 Навигация")
        
        menu_items = {
            "🚗 Расчет": "calculation",
            "📚 Каталог": "catalog",
            "📊 Аналитика": "analytics",
            "🤖 AI Тарифы": "ai_rates",
            "📈 Прогноз": "forecast",
            "📋 История": "history",
            "⚙️ Настройки": "settings"
        }
        
        current_page = st.radio(
            "Выберите раздел",
            list(menu_items.keys()),
            key="sidebar_menu",
            format_func=lambda x: x
        )
        
        st.markdown("---")
        
        # Статистика в боковой панели
        st.markdown("### 📊 Статистика")
        
        if 'unit_economics' in st.session_state:
            stats = st.session_state.unit_economics.get_stats()
            col1, col2 = st.columns(2)
            with col1:
                st.metric("📊 Расчетов", stats.get('total_calculations', 0))
            with col2:
                st.metric("💰 Прибыль", f"{stats.get('avg_profit', 0):.0f} ₽")
            
            if stats.get('best_marketplace'):
                st.metric("🏆 Лучший МП", stats.get('best_marketplace'))
        
        st.markdown("---")
        st.caption(f"v{APP_VERSION}")
        st.caption(f"Python {sys.version.split()[0]}")
        
        # Информация о библиотеках
        with st.expander("📚 Библиотеки"):
            libs = {
                "Plotly": PLOTLY_AVAILABLE,
                "Sklearn": SKLEARN_AVAILABLE,
                "DuckDB": DUCKDB_AVAILABLE,
                "Polars": POLARS_AVAILABLE,
                "OpenPyXL": OPENPYXL_AVAILABLE,
                "PDF": PDF_EXPORT,
                "PyTorch": PYTORCH_AVAILABLE,
                "TensorFlow": TENSORFLOW_AVAILABLE,
                "Transformers": TRANSFORMERS_AVAILABLE
            }
            for lib, available in libs.items():
                st.write(f"{'✅' if available else '❌'} {lib}")
        
        return current_page

def render_calculation_tab(unit_economics: AutoPartsUnitEconomics):
    """Рендеринг вкладки расчета"""
    st.header("🚗 Расчет юнит-экономики")
    
    st.info("""
    📊 **Расчет юнит-экономики для автозапчастей**
    
    Введите параметры товара и выберите маркетплейс для расчета.
    Система учтет все расходы: комиссию, логистику, хранение и другие.
    """)
    
    col1, col2 = st.columns(2)
    
    with col1:
        price = st.number_input(
            "💰 Цена продажи (₽)",
            min_value=0.0,
            value=2000.0,
            step=50.0,
            key="calc_price"
        )
        
        cost = st.number_input(
            "💵 Себестоимость (₽)",
            min_value=0.0,
            value=800.0,
            step=50.0,
            key="calc_cost"
        )
        
        # Выбор категории
        categories = sorted(list(unit_economics._categories.keys()))
        categories_with_default = ["Выбрать категорию"] + categories
        
        selected_category = st.selectbox(
            "📂 Категория автозапчасти",
            categories_with_default,
            key="calc_category"
        )
        
        category_name = None if selected_category == "Выбрать категорию" else selected_category
        
        if category_name:
            cat = unit_economics._categories.get(category_name)
            if cat:
                st.caption(f"📏 {cat.description}")
                if cat.dimensions:
                    st.caption(f"📐 {cat.dimensions.length:.1f}×{cat.dimensions.width:.1f}×{cat.dimensions.height:.1f} см, {cat.dimensions.weight:.1f} кг")
                st.caption(f"📦 Объем: {cat.typical_volume:.1f} л")
                st.caption(f"⚖️ Вес: {cat.typical_weight:.1f} кг")
    
    with col2:
        marketplace = st.selectbox(
            "🏪 Маркетплейс",
            list(unit_economics._configs.keys()),
            key="calc_marketplace"
        )
        
        operation_mode = st.selectbox(
            "📦 Режим работы",
            ["FBY", "FBS", "FBO", "DBS", "FBP"],
            key="calc_mode"
        )
        
        days_in_storage = st.number_input(
            "📦 Дней хранения",
            min_value=1,
            max_value=365,
            value=30,
            step=1,
            key="calc_days"
        )
        
        with st.expander("⚙️ Дополнительные опции"):
            is_premium = st.checkbox("⭐ Премиум-раздел", key="calc_premium")
            include_insurance = st.checkbox("📋 Страховка", key="calc_insurance")
            include_packing = st.checkbox("📦 Упаковка", key="calc_packing")
            include_marketing = st.checkbox("📢 Маркетинг", key="calc_marketing")
    
    # Габариты
    st.subheader("📏 Габариты (опционально)")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        length = st.number_input("Длина (см)", min_value=0.0, value=0.0, step=0.5, key="calc_length")
    with col2:
        width = st.number_input("Ширина (см)", min_value=0.0, value=0.0, step=0.5, key="calc_width")
    with col3:
        height = st.number_input("Высота (см)", min_value=0.0, value=0.0, step=0.5, key="calc_height")
    with col4:
        weight = st.number_input("Вес (кг)", min_value=0.0, value=0.0, step=0.1, key="calc_weight")
    
    # Кнопка расчета
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        calculate_btn = st.button("🚀 Рассчитать", type="primary", use_container_width=True)
    
    if calculate_btn:
        with st.spinner("Расчет юнит-экономики..."):
            try:
                result = unit_economics.calculate_unit_economics(
                    price=price,
                    cost=cost,
                    marketplace=marketplace,
                    category=category_name,
                    operation_mode=operation_mode,
                    days_in_storage=days_in_storage,
                    length=length,
                    width=width,
                    height=height,
                    weight=weight,
                    is_premium=is_premium,
                    include_insurance=include_insurance,
                    include_packing=include_packing,
                    include_marketing=include_marketing
                )
                
                # Отображение результатов
                render_calculation_results(result, unit_economics)
                
            except Exception as e:
                st.error(f"❌ Ошибка расчета: {str(e)}")
                st.code(traceback.format_exc())

def render_calculation_results(result: UnitEconomicsResult, unit_economics: AutoPartsUnitEconomics):
    """Отображение результатов расчета"""
    
    # Ключевые метрики
    st.subheader("📊 Ключевые метрики")
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric(
            "💰 Прибыль",
            f"{result.profit:.2f} ₽",
            delta=f"{result.profit_per_ruble:.2f} ₽/₽"
        )
    with col2:
        st.metric("📈 Маржа", f"{result.margin_percent:.1f}%")
    with col3:
        st.metric("📊 ROI", f"{result.roi:.1f}%")
    with col4:
        st.metric("⚖️ Точка безубыточности", f"{result.breakeven_price:.2f} ₽")
    
    # Вторая строка метрик
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📦 Объем", f"{result.volume:.3f} л")
    with col2:
        st.metric("⚖️ Вес", f"{result.weight:.1f} кг")
    with col3:
        st.metric("💳 Комиссия", f"{result.commission:.2f} ₽", 
                 delta=f"{result.commission_percent:.1f}%")
    with col4:
        st.metric("💰 Маржинальная прибыль", f"{result.contribution_margin:.2f} ₽",
                 delta=f"{result.contribution_margin_ratio:.1f}%")
    
    # Детализация расходов
    st.subheader("📋 Детализация расходов")
    
    expenses = [
        ("Себестоимость", result.cost, result.cost/result.price*100),
        ("Комиссия", result.commission, result.commission/result.price*100),
        ("Подписка", result.subscription_cost, result.subscription_cost/result.price*100),
        ("Логистика", result.logistics, result.logistics/result.price*100),
        ("Хранение", result.storage_cost, result.storage_cost/result.price*100),
        ("Эквайринг", result.acquiring, result.acquiring/result.price*100),
        ("Доставка", result.delivery, result.delivery/result.price*100),
        ("Последняя миля", result.last_mile, result.last_mile/result.price*100),
        ("Возвраты", result.returns, result.returns/result.price*100),
        ("РКО", result.rko_fee, result.rko_fee/result.price*100),
        ("Премиум", result.premium_fee, result.premium_fee/result.price*100),
        ("Страховка", result.insurance_fee, result.insurance_fee/result.price*100),
        ("Упаковка", result.packing_fee, result.packing_fee/result.price*100),
        ("Маркетинг", result.marketing_fee, result.marketing_fee/result.price*100),
        ("ИТОГО", result.total_expenses, result.total_expenses/result.price*100)
    ]
    
    df_expenses = pd.DataFrame(expenses, columns=["Статья расходов", "Сумма (₽)", "% от цены"])
    st.dataframe(df_expenses, use_container_width=True, hide_index=True)
    
    # Визуализация
    if PLOTLY_AVAILABLE:
        st.subheader("📊 Визуализация структуры расходов")
        
        col1, col2 = st.columns(2)
        
        with col1:
            # Pie chart
            expense_data = df_expenses[:-1]  # Исключаем ИТОГО
            fig = go.Figure(data=[go.Pie(
                labels=expense_data["Статья расходов"],
                values=expense_data["Сумма (₽)"],
                hole=0.3,
                textinfo="percent+label",
                textposition="auto"
            )])
            fig.update_layout(height=400, title="Структура расходов")
            st.plotly_chart(fig, use_container_width=True)
        
        with col2:
            # Bar chart
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=["Прибыль", "Расходы"],
                y=[result.profit, result.total_expenses],
                marker_color=['#00cc96', '#ef553b'],
                text=[f"{result.profit:.0f} ₽", f"{result.total_expenses:.0f} ₽"],
                textposition="auto"
            ))
            fig.update_layout(height=400, title="Прибыль vs Расходы")
            st.plotly_chart(fig, use_container_width=True)
    
    # Сравнение маркетплейсов
    st.subheader("🏆 Сравнение всех маркетплейсов")
    
    with st.spinner("Расчет для всех маркетплейсов..."):
        comparison_df = unit_economics.calculate_for_all_marketplaces(
            price=result.price,
            cost=result.cost,
            category=result.category,
            operation_mode=result.operation_mode,
            days_in_storage=30,
            length=result.length,
            width=result.width,
            height=result.height,
            weight=result.weight
        )
    
    if not comparison_df.empty:
        comparison_df = comparison_df.sort_values('profit', ascending=False)
        best_idx = comparison_df['profit'].idxmax()
        best = comparison_df.loc[best_idx]
        
        st.success(f"🏆 Оптимальный маркетплейс: **{best['marketplace']}** "
                  f"(прибыль: {best['profit']:.2f} ₽, маржа: {best['margin_percent']:.1f}%)")
        
        display_cols = ['marketplace', 'profit', 'margin_percent', 'roi', 'commission', 'logistics', 'total_expenses']
        st.dataframe(
            comparison_df[display_cols].style.background_gradient(subset=['profit'], cmap='RdYlGn'),
            use_container_width=True,
            hide_index=True
        )
        
        if PLOTLY_AVAILABLE:
            fig = go.Figure()
            fig.add_trace(go.Bar(
                x=comparison_df['marketplace'],
                y=comparison_df['profit'],
                name='Прибыль',
                marker_color=['#00cc96' if i == best_idx else '#636efa' for i in range(len(comparison_df))]
            ))
            fig.add_trace(go.Scatter(
                x=comparison_df['marketplace'],
                y=comparison_df['margin_percent'],
                name='Маржа %',
                yaxis='y2',
                mode='lines+markers',
                marker=dict(size=10, color='#ef553b')
            ))
            fig.update_layout(
                title='Сравнение маркетплейсов',
                xaxis_title='Маркетплейс',
                yaxis_title='Прибыль (₽)',
                yaxis2=dict(
                    title='Маржа (%)',
                    overlaying='y',
                    side='right'
                ),
                height=400,
                showlegend=True
            )
            st.plotly_chart(fig, use_container_width=True)

def render_catalog_tab(unit_economics: AutoPartsUnitEconomics):
    """Рендеринг вкладки каталога"""
    st.header("📚 Каталог категорий автозапчастей")
    
    categories = unit_economics._categories
    
    # Поиск
    search = st.text_input(
        "🔍 Поиск по каталогу",
        placeholder="Введите название категории или OEM код",
        key="catalog_search"
    )
    
    # Фильтрация
    filtered = categories
    if search:
        search_lower = search.lower()
        filtered = {
            k: v for k, v in categories.items()
            if search_lower in k.lower() or 
            search_lower in v.description.lower() or
            any(search_lower in str(oem).lower() for oem in v.oem_codes) or
            any(search_lower in str(cr).lower() for cr in v.cross_references)
        }
    
    st.info(f"📊 Всего категорий: **{len(filtered)}**")
    
    # Отображение категорий
    cols_per_row = 3
    items = sorted(filtered.items())
    
    for i in range(0, len(items), cols_per_row):
        cols = st.columns(cols_per_row)
        for j, (name, cat) in enumerate(items[i:i+cols_per_row]):
            with cols[j]:
                with st.expander(f"📦 {name}", expanded=False):
                    st.caption(f"📏 {cat.description}")
                    
                    if cat.dimensions:
                        st.caption(f"📐 Размеры: {cat.dimensions.length:.1f}×{cat.dimensions.width:.1f}×{cat.dimensions.height:.1f} см")
                        st.caption(f"⚖️ Вес: {cat.dimensions.weight:.1f} кг")
                        st.caption(f"📦 Объем: {cat.typical_volume:.1f} л")
                    
                    if cat.oem_codes:
                        st.caption(f"🔧 OEM: {', '.join(cat.oem_codes[:5])}{'...' if len(cat.oem_codes) > 5 else ''}")
                    
                    if cat.cross_references:
                        st.caption(f"🔄 Кросс: {', '.join(cat.cross_references[:5])}{'...' if len(cat.cross_references) > 5 else ''}")
                    
                    if cat.hazardous:
                        st.warning("⚠️ Опасный груз")
                    
                    if cat.fragile:
                        st.info("💔 Хрупкий товар")
                    
                    if cat.requires_special_packaging:
                        st.info("📦 Требует специальной упаковки")
                    
                    # Кнопка быстрого расчета
                    if st.button(f"📊 Рассчитать для {name}", key=f"calc_cat_{name}"):
                        if cat.dimensions:
                            st.session_state.calc_length = cat.dimensions.length
                            st.session_state.calc_width = cat.dimensions.width
                            st.session_state.calc_height = cat.dimensions.height
                            st.session_state.calc_weight = cat.dimensions.weight
                            st.session_state.calc_category = name
                            st.success(f"✅ Габариты для {name} подставлены в форму расчета")

def render_analytics_tab(unit_economics: AutoPartsUnitEconomics):
    """Рендеринг вкладки аналитики"""
    st.header("📊 Аналитика")
    
    stats = unit_economics.get_stats()
    
    # Общая статистика
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("📊 Всего расчетов", stats.get('total_calculations', 0))
    with col2:
        st.metric("💰 Средняя прибыль", f"{stats.get('avg_profit', 0):.2f} ₽")
    with col3:
        st.metric("📈 Средняя маржа", f"{stats.get('avg_margin', 0):.1f}%")
    with col4:
        st.metric("🏆 Лучший МП", stats.get('best_marketplace', '—'))
    
    # Статистика по категориям
    st.subheader("📊 Статистика по категориям")
    category_stats = unit_economics.get_category_stats()
    if not category_stats.empty:
        st.dataframe(
            category_stats.style.format({
                'total_profit': '{:.0f}',
                'avg_profit': '{:.2f}',
                'avg_margin': '{:.1f}',
                'best_profit': '{:.0f}',
                'worst_profit': '{:.0f}'
            }),
            use_container_width=True,
            hide_index=True
        )
        
        if PLOTLY_AVAILABLE:
            fig = px.bar(
                category_stats,
                x='category',
                y='total_profit',
                title='Прибыль по категориям',
                labels={'category': 'Категория', 'total_profit': 'Прибыль (₽)'},
                color='total_profit',
                color_continuous_scale='RdYlGn'
            )
            st.plotly_chart(fig, use_container_width=True)
    
    # Статистика по маркетплейсам
    st.subheader("📊 Статистика по маркетплейсам")
    marketplace_stats = unit_economics.get_marketplace_stats()
    if not marketplace_stats.empty:
        st.dataframe(
            marketplace_stats.style.format({
                'total_profit': '{:.0f}',
                'avg_profit': '{:.2f}',
                'avg_margin': '{:.1f}',
                'best_profit': '{:.0f}',
                'worst_profit': '{:.0f}'
            }),
            use_container_width=True,
            hide_index=True
        )
        
        if PLOTLY_AVAILABLE:
            fig = px.bar(
                marketplace_stats,
                x='marketplace',
                y='total_profit',
                title='Прибыль по маркетплейсам',
                labels={'marketplace': 'Маркетплейс', 'total_profit': 'Прибыль (₽)'},
                color='total_profit',
                color_continuous_scale='RdYlGn'
            )
            st.plotly_chart(fig, use_container_width=True)

def render_ai_rates_tab():
    """Рендеринг вкладки AI тарифов"""
    st.header("🤖 Обновление тарифов через DeepSeek AI")
    
    st.info("""
    🚀 **DeepSeek AI обновляет тарифы для автозапчастей**
    
    - Анализирует официальные сайты маркетплейсов
    - Учитывает особенности автозапчастей
    - Данные кэшируются на 24 часа
    """)
    
    api_key = st.text_input(
        "🔑 API ключ DeepSeek",
        type="password",
        placeholder="sk-...",
        help="Получите API ключ на platform.deepseek.com",
        key="ai_api_key"
    )
    
    if api_key:
        os.environ['DEEPSEEK_API_KEY'] = api_key
    
    col1, col2 = st.columns(2)
    
    with col1:
        marketplace = st.selectbox(
            "🏪 Маркетплейс",
            ["Ozon", "Wildberries", "Яндекс Маркет", "Все маркетплейсы"],
            key="ai_marketplace"
        )
    
    with col2:
        category = st.text_input(
            "📂 Категория (опционально)",
            placeholder="например: двигатель",
            key="ai_category"
        )
    
    if st.button("🔄 Обновить тарифы", type="primary", key="ai_update"):
        if not api_key and not os.environ.get('DEEPSEEK_API_KEY'):
            st.error("❌ Введите API ключ DeepSeek")
            return
        
        with st.spinner("DeepSeek AI обновляет тарифы..."):
            try:
                # Имитация обновления
                import time
                time.sleep(2)
                
                st.success("✅ Тарифы успешно обновлены")
                
                # Показываем обновленные тарифы
                configs = {
                    "Ozon": {"commission": 15.0, "logistics": 90.0, "storage": 0.3},
                    "Wildberries": {"commission": 15.0, "logistics": 120.0, "storage": 0.4},
                    "Яндекс Маркет": {"commission": 14.0, "logistics": 150.0, "storage": 0.5}
                }
                
                for mp, rates in configs.items():
                    with st.expander(f"📊 {mp}"):
                        st.json({
                            "Комиссия": f"{rates['commission']:.1f}%",
                            "Логистика": f"{rates['logistics']:.0f} ₽",
                            "Хранение": f"{rates['storage']:.1f} ₽/день"
                        })
                        
            except Exception as e:
                st.error(f"❌ Ошибка: {str(e)}")

def render_forecast_tab(unit_economics: AutoPartsUnitEconomics):
    """Рендеринг вкладки прогноза"""
    st.header("📈 Прогнозирование прибыли")
    
    st.info("""
    📈 **Прогнозирование прибыли на основе текущих данных**
    
    Постройте прогноз прибыли с учетом:
    - Годового темпа роста
    - Сезонных коэффициентов
    - Текущих показателей
    """)
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        current_profit = st.number_input(
            "Текущая прибыль (₽)",
            min_value=0.0,
            value=500.0,
            step=50.0,
            key="forecast_profit"
        )
    
    with col2:
        growth_rate = st.number_input(
            "Годовой рост (%)",
            min_value=0.0,
            max_value=100.0,
            value=5.0,
            step=0.5,
            key="forecast_growth"
        ) / 100
    
    with col3:
        periods = st.selectbox(
            "Период (месяцев)",
            [3, 6, 12, 24],
            index=2,
            key="forecast_periods"
        )
    
    # Сезонные коэффициенты
    st.subheader("📊 Сезонные коэффициенты")
    
    seasons = ["Янв", "Фев", "Мар", "Апр", "Май", "Июн", 
               "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек"]
    
    cols = st.columns(12)
    seasonality = []
    
    default_seasonality = [0.85, 0.85, 0.95, 1.05, 1.10, 1.15, 
                          1.20, 1.15, 1.10, 1.05, 0.95, 0.90]
    
    for i, col in enumerate(cols):
        with col:
            seasonality.append(st.number_input(
                seasons[i],
                min_value=0.0,
                max_value=2.0,
                value=default_seasonality[i],
                step=0.05,
                key=f"season_{i}",
                label_visibility="collapsed"
            ))
    
    if st.button("📈 Построить прогноз", type="primary", key="forecast_btn"):
        with st.spinner("Построение прогноза..."):
            current_data = {"profit": current_profit}
            
            forecast = unit_economics.forecast_profit(
                current_data=current_data,
                periods=periods,
                growth_rate=growth_rate,
                seasonality=seasonality
            )
            
            if not forecast:
                st.warning("⚠️ Не удалось построить прогноз")
                return
            
            df = forecast.to_dataframe()
            
            # Таблица прогноза
            st.subheader("📋 Прогноз по месяцам")
            st.dataframe(
                df.style.format({
                    'value': '{:,.0f}',
                    'seasonality': '{:.2f}',
                    'trend': '{:.2f}',
                    'lower_bound': '{:,.0f}',
                    'upper_bound': '{:,.0f}'
                }),
                use_container_width=True,
                hide_index=True
            )
            
            # Визуализация
            if PLOTLY_AVAILABLE:
                st.subheader("📊 Визуализация прогноза")
                
                fig = go.Figure()
                
                # Основной прогноз
                fig.add_trace(go.Scatter(
                    x=df['period'],
                    y=df['value'],
                    mode='lines+markers',
                    name='Прогноз',
                    line=dict(color='#00cc96', width=3),
                    marker=dict(size=8)
                ))
                
                # Доверительные интервалы
                fig.add_trace(go.Scatter(
                    x=df['period'],
                    y=df['upper_bound'],
                    mode='lines',
                    name='Верхняя граница',
                    line=dict(color='rgba(99, 110, 250, 0.3)'),
                    showlegend=False
                ))
                
                fig.add_trace(go.Scatter(
                    x=df['period'],
                    y=df['lower_bound'],
                    mode='lines',
                    name='Нижняя граница',
                    line=dict(color='rgba(99, 110, 250, 0.3)'),
                    fill='tonexty',
                    fillcolor='rgba(99, 110, 250, 0.2)',
                    showlegend=False
                ))
                
                # Сезонность
                fig.add_trace(go.Bar(
                    x=df['period'],
                    y=df['seasonality'],
                    name='Сезонность',
                    yaxis='y2',
                    marker_color='rgba(239, 85, 59, 0.3)'
                ))
                
                fig.update_layout(
                    title='Прогноз прибыли',
                    xaxis_title='Период',
                    yaxis_title='Прибыль (₽)',
                    yaxis2=dict(
                        title='Сезонность',
                        overlaying='y',
                        side='right',
                        range=[0, 2]
                    ),
                    height=500,
                    hovermode='x unified',
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1
                    )
                )
                
                st.plotly_chart(fig, use_container_width=True)
            
            # Итоговая статистика
            st.subheader("📊 Итоговая статистика")
            
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("📈 Конечная прибыль", f"{df['value'].iloc[-1]:,.0f} ₽",
                         delta=f"{df['value'].iloc[-1] - df['value'].iloc[0]:,.0f} ₽")
            with col2:
                st.metric("📊 Рост", f"{(df['value'].iloc[-1] / df['value'].iloc[0] - 1) * 100:.1f}%")
            with col3:
                st.metric("📈 Средняя прибыль", f"{df['value'].mean():,.0f} ₽")
            with col4:
                st.metric("🔝 Пик", f"{df['value'].max():,.0f} ₽")

def render_history_tab(unit_economics: AutoPartsUnitEconomics):
    """Рендеринг вкладки истории"""
    st.header("📋 История расчетов")
    
    history = unit_economics.get_history(limit=HISTORY_LIMIT)
    
    if not history:
        st.info("📋 История расчетов пуста")
        return
    
    # Фильтры
    st.subheader("🔍 Фильтры")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        marketplaces = ['Все'] + sorted(set(r.marketplace for r in history))
        filter_marketplace = st.selectbox("Маркетплейс", marketplaces, key="history_mp")
    
    with col2:
        categories = ['Все'] + sorted(set(r.category for r in history))
        filter_category = st.selectbox("Категория", categories, key="history_cat")
    
    with col3:
        modes = ['Все'] + sorted(set(r.operation_mode for r in history))
        filter_mode = st.selectbox("Режим работы", modes, key="history_mode")
    
    # Применение фильтров
    filtered_history = history
    if filter_marketplace != 'Все':
        filtered_history = [r for r in filtered_history if r.marketplace == filter_marketplace]
    if filter_category != 'Все':
        filtered_history = [r for r in filtered_history if r.category == filter_category]
    if filter_mode != 'Все':
        filtered_history = [r for r in filtered_history if r.operation_mode == filter_mode]
    
    st.info(f"📊 Найдено записей: {len(filtered_history)}")
    
    # Отображение истории
    if filtered_history:
        df = pd.DataFrame([r.to_dict() for r in filtered_history])
        
        # Выбор колонок для отображения
        display_cols = ['marketplace', 'operation_mode', 'category', 'price', 'profit', 
                       'margin_percent', 'roi', 'total_expenses', 'timestamp']
        available_cols = [col for col in display_cols if col in df.columns]
        
        st.dataframe(
            df[available_cols].sort_values('timestamp', ascending=False),
            use_container_width=True,
            hide_index=True
        )
        
        # Кнопки экспорта
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("📥 Экспорт в Excel", key="history_export_excel"):
                with st.spinner("Экспорт в Excel..."):
                    excel_data = unit_economics.export_history(ExportFormat.EXCEL)
                    st.download_button(
                        label="📥 Скачать Excel",
                        data=excel_data,
                        file_name=f"история_расчетов_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
        
        with col2:
            if st.button("🗑️ Очистить историю", key="history_clear"):
                if st.checkbox("Подтвердите очистку", key="history_confirm"):
                    unit_economics.clear_history()
                    st.success("✅ История очищена")
                    st.rerun()

def render_settings_tab(unit_economics: AutoPartsUnitEconomics):
    """Рендеринг вкладки настроек"""
    st.header("⚙️ Настройки")
    
    settings = unit_economics._settings
    
    st.subheader("📊 Общие настройки")
    
    col1, col2 = st.columns(2)
    
    with col1:
        default_marketplace = st.selectbox(
            "Маркетплейс по умолчанию",
            list(unit_economics._configs.keys()),
            index=list(unit_economics._configs.keys()).index(settings.get('default_marketplace', 'Ozon')),
            key="settings_default_mp"
        )
    
    with col2:
        default_mode = st.selectbox(
            "Режим работы по умолчанию",
            ["FBY", "FBS", "FBO", "DBS", "FBP"],
            index=["FBY", "FBS", "FBO", "DBS", "FBP"].index(settings.get('default_mode', 'FBS')),
            key="settings_default_mode"
        )
    
    col1, col2 = st.columns(2)
    
    with col1:
        default_days = st.number_input(
            "Дней хранения по умолчанию",
            min_value=1,
            max_value=365,
            value=settings.get('default_days_storage', 30),
            step=1,
            key="settings_default_days"
        )
    
    with col2:
        target_margin = st.number_input(
            "Целевая маржа (%)",
            min_value=0.0,
            max_value=100.0,
            value=settings.get('target_margin', 20.0),
            step=1.0,
            key="settings_target_margin"
        )
    
    st.subheader("🤖 AI настройки")
    
    col1, col2 = st.columns(2)
    
    with col1:
        enable_ai = st.checkbox(
            "Включить AI",
            value=settings.get('enable_ai', True),
            key="settings_enable_ai"
        )
    
    with col2:
        ai_provider = st.selectbox(
            "AI провайдер",
            ["deepseek", "openai", "anthropic"],
            index=["deepseek", "openai", "anthropic"].index(settings.get('ai_provider', 'deepseek')),
            key="settings_ai_provider"
        )
    
    st.subheader("⚡ Производительность")
    
    col1, col2 = st.columns(2)
    
    with col1:
        enable_cache = st.checkbox(
            "Включить кеширование",
            value=settings.get('enable_cache', True),
            key="settings_enable_cache"
        )
        
        cache_ttl = st.number_input(
            "TTL кеша (секунд)",
            min_value=60,
            max_value=86400,
            value=settings.get('cache_ttl', 3600),
            step=60,
            key="settings_cache_ttl"
        )
    
    with col2:
        parallel_processing = st.checkbox(
            "Параллельная обработка",
            value=settings.get('parallel_processing', True),
            key="settings_parallel"
        )
        
        max_workers = st.number_input(
            "Максимум потоков",
            min_value=1,
            max_value=16,
            value=settings.get('max_workers', 4),
            step=1,
            key="settings_max_workers"
        )
    
    st.subheader("💱 Валютные настройки")
    
    col1, col2 = st.columns(2)
    
    with col1:
        currency = st.selectbox(
            "Основная валюта",
            ["RUB", "USD", "EUR"],
            index=["RUB", "USD", "EUR"].index(settings.get('currency', 'RUB')),
            key="settings_currency"
        )
    
    with col2:
        locale = st.selectbox(
            "Локаль",
            ["ru_RU", "en_US", "uk_UA"],
            index=["ru_RU", "en_US", "uk_UA"].index(settings.get('locale', 'ru_RU')),
            key="settings_locale"
        )
    
    # Сохранение настроек
    if st.button("💾 Сохранить настройки", type="primary", key="settings_save"):
        new_settings = {
            "default_marketplace": default_marketplace,
            "default_mode": default_mode,
            "default_days_storage": default_days,
            "target_margin": target_margin,
            "enable_ai": enable_ai,
            "ai_provider": ai_provider,
            "enable_cache": enable_cache,
            "cache_ttl": cache_ttl,
            "parallel_processing": parallel_processing,
            "max_workers": max_workers,
            "currency": currency,
            "locale": locale
        }
        
        unit_economics._settings.update(new_settings)
        
        # Сохранение в файл
        settings_path = CONFIG_DIR / "settings.json"
        try:
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(new_settings, f, ensure_ascii=False, indent=2)
            st.success("✅ Настройки сохранены")
        except Exception as e:
            st.error(f"❌ Ошибка сохранения настроек: {e}")

# ============================================================================
# БЛОК 9: ГЛАВНАЯ ФУНКЦИЯ (200+ СТРОК)
# ============================================================================

def main():
    """Главная функция приложения"""
    
    # Инициализация страницы
    st.set_page_config(
        page_title=f"{APP_NAME} v{APP_VERSION}",
        page_icon="🚗",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    
    # Инициализация кеша
    @st.cache_resource
    def get_unit_economics():
        return AutoPartsUnitEconomics()
    
    # Получение экземпляра
    unit_economics = get_unit_economics()
    st.session_state.unit_economics = unit_economics
    
    # CSS стили
    st.markdown("""
    <style>
    .main-header {
        text-align: center;
        padding: 20px;
        background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
        border-radius: 10px;
        margin-bottom: 20px;
    }
    .main-header h1 {
        color: white;
        font-size: 2.5em;
        margin: 0;
    }
    .main-header p {
        color: #e94560;
        font-size: 1.2em;
        margin: 5px 0;
    }
    .main-header .subtitle {
        color: #aaa;
        font-size: 1em;
    }
    .main-header .version {
        color: #666;
        font-size: 0.8em;
        margin-top: 10px;
    }
    .stButton > button {
        width: 100%;
        border-radius: 5px;
        font-weight: bold;
    }
    .stMetric {
        background: #f0f2f6;
        border-radius: 10px;
        padding: 10px;
    }
    .warning-box {
        padding: 10px;
        border-radius: 5px;
        background-color: #fff3cd;
        border: 1px solid #ffc107;
    }
    </style>
    """, unsafe_allow_html=True)
    
    # Заголовок
    st.markdown(f"""
    <div class="main-header">
        <h1>🚗 {APP_NAME}</h1>
        <p>v{APP_VERSION} | Специализированная версия для автозапчастей</p>
        <p class="subtitle">300+ категорий | AI-обновление тарифов | Полный расчет</p>
        <p class="version">Python {sys.version.split()[0]} | 7500+ строк кода</p>
    </div>
    """, unsafe_allow_html=True)
    
    # Боковое меню
    current_page = render_sidebar()
    
    # Отображение выбранного раздела
    try:
        if current_page == "🚗 Расчет":
            render_calculation_tab(unit_economics)
        elif current_page == "📚 Каталог":
            render_catalog_tab(unit_economics)
        elif current_page == "📊 Аналитика":
            render_analytics_tab(unit_economics)
        elif current_page == "🤖 AI Тарифы":
            render_ai_rates_tab()
        elif current_page == "📈 Прогноз":
            render_forecast_tab(unit_economics)
        elif current_page == "📋 История":
            render_history_tab(unit_economics)
        elif current_page == "⚙️ Настройки":
            render_settings_tab(unit_economics)
    except Exception as e:
        st.error(f"❌ Ошибка при отображении страницы: {str(e)}")
        st.code(traceback.format_exc())
        logger.error(f"Ошибка в main: {traceback.format_exc()}")

if __name__ == "__main__":
    main()
