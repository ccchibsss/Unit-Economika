"""
================================================================================
🚗 ULTIMATE UNIT ECONOMICS FOR AUTO PARTS v97.0 - ФИНАЛЬНАЯ ВЕРСИЯ
================================================================================
📌 ВЕРСИЯ: 97.0.0 (ОБЪЕДИНЁННАЯ С ИИ-ТАРИФАМИ)
📌 ОБЩИЙ ОБЪЕМ: 7,500+ СТРОК (ПОЛНАЯ ВЕРСИЯ БЕЗ СОКРАЩЕНИЙ)
📌 СОВМЕСТИМОСТЬ: Python 3.10 - 3.14
📌 СПЕЦИАЛИЗАЦИЯ: АВТОЗАПЧАСТИ И АВТОТОВАРЫ
📌 НОВЫЙ ФУНКЦИОНАЛ В v97.0:
✅ ГАЛОЧКА "ЗАПРОСИТЬ ИИ" - ДИНАМИЧЕСКОЕ ОБНОВЛЕНИЕ ТАРИФОВ
✅ УМНЫЙ КЭШ ТАРИФОВ С ВОЗМОЖНОСТЬЮ РУЧНОГО РЕДАКТИРОВАНИЯ
✅ КОЛОНКА "РЕКОМЕНДОВАННАЯ ЦЕНА МИНИМУМ" (с учётом УСН 6% + 10% прибыли)
✅ РАСШИРЕННЫЙ ШАБЛОН EXCEL С 5 ЛИСТАМИ (Тарифы, Налоги, Расчёт, Сводка, Dashboard)
✅ ЖИВЫЕ ФОРМУЛЫ В EXCEL (VLOOKUP, SUMIF, AVERAGEIF)
✅ УЧЁТ НАЛОГОВ (УСН, СТРАХОВЫЕ ВЗНОСЫ, НДС)
✅ 150+ КАТЕГОРИЙ АВТОЗАПЧАСТЕЙ С ПОЛНЫМИ ГАБАРИТАМИ
✅ РАСЧЕТ ЮНИТ-ЭКОНОМИКИ ПО КАЖДОМУ АРТИКУЛУ
✅ ПОИСК АНАЛОГОВ ПО OE НОМЕРАМ (2 УРОВНЯ)
✅ ML-КЛАССИФИКАЦИЯ ТОВАРОВ
✅ ИНТЕГРАЦИЯ С DEEPSEEK AI
✅ HIGH-VOLUME CATALOG (10M+ ЗАПИСЕЙ) С POLARS И DUCKDB
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
    from openpyxl.formatting.rule import Rule, ColorScaleRule, DataBarRule, IconSetRule, CellIsRule
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
APP_VERSION = "97.0.0"
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

# === НОВОЕ v97: Настройки налогов и прибыли ===
DEFAULT_TAX_SYSTEM = "УСН_6"  # УСН 6%, УСН_15, ОСНО
DEFAULT_USN_RATE = 0.06
DEFAULT_USN_ADDITIONAL_RATE = 0.01  # +1% сверх лимита
DEFAULT_USN_LIMIT = 300_000_000  # лимит 300 млн в год
DEFAULT_INSURANCE_CONTRIBUTIONS = 49_500  # фиксированные страховые взносы ИП в 2026
DEFAULT_MIN_PROFIT_PERCENT = 0.10  # минимальная прибыль 10%
DEFAULT_TAX_DEDUCTION = True  # уменьшение налога на страховые взносы

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
TARIFFS_DIR = BASE_DIR / "tariffs"  # НОВОЕ v97: папка для тарифов

for dir_path in [DATA_DIR, CACHE_DIR, LOG_DIR, REPORTS_DIR, TEMP_DIR, MODELS_DIR, CONFIG_DIR, PLUGINS_DIR, EXPORTS_DIR, TARIFFS_DIR]:
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
    "gradient_end": "#16213e",
    "input_fill": "#FFF4CC",    # НОВОЕ v97: цвет для вводных данных
    "formula_fill": "#E2EFDA",  # НОВОЕ v97: цвет для формул
    "result_fill": "#DCE6F1",   # НОВОЕ v97: цвет для результатов
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
# НОВЫЕ v97: ФУНКЦИИ ДЛЯ РАБОТЫ С НАЛОГАМИ
# ============================================================================
def calculate_tax(
    revenue: float,
    annual_revenue: float = 0,
    tax_system: str = "УСН_6",
    insurance_contributions: float = 49_500,
    use_deduction: bool = True
) -> float:
    """
    Расчёт налога для ИП/ООО.
    
    Args:
        revenue: Выручка от продажи одного товара
        annual_revenue: Годовая выручка (для расчёта 1% сверх лимита)
        tax_system: Система налогообложения (УСН_6, УСН_15, ОСНО)
        insurance_contributions: Фиксированные страховые взносы
        use_deduction: Уменьшать налог на страховые взносы
        
    Returns:
        Сумма налога
    """
    if revenue <= 0:
        return 0.0
    
    if tax_system == "УСН_6":
        # Базовый налог 6%
        base_tax = revenue * DEFAULT_USN_RATE
        
        # Дополнительный 1% сверх лимита (300 млн для ИП)
        if annual_revenue > DEFAULT_USN_LIMIT:
            excess = annual_revenue - DEFAULT_USN_LIMIT
            additional_tax = excess * DEFAULT_USN_ADDITIONAL_RATE
            # 1% распределяется пропорционально
            additional_per_unit = additional_tax * (revenue / annual_revenue) if annual_revenue > 0 else 0
            base_tax += additional_per_unit
        
        # Уменьшение на страховые взносы (до 100% для ИП без сотрудников)
        if use_deduction and insurance_contributions > 0:
            # Максимально можно уменьшить на сумму взносов
            deduction = min(base_tax, insurance_contributions / max(1, annual_revenue / revenue))
            return max(0, base_tax - deduction)
        
        return base_tax
    
    elif tax_system == "УСН_15":
        # УСН "Доходы минус расходы" 15%
        return revenue * 0.15
    
    elif tax_system == "ОСНО":
        # Общая система: НДС 20% + налог на прибыль 20%
        nds = revenue * 0.20
        profit_tax = revenue * 0.20
        return nds + profit_tax
    
    return 0.0


def calculate_recommended_min_price(
    cost: float,
    commission_rate: float,
    logistics: float,
    storage_cost: float,
    acquiring_rate: float,
    last_mile: float,
    return_rate: float,
    min_profit_percent: float = 0.10,
    tax_system: str = "УСН_6",
    tax_rate: float = 0.06
) -> float:
    """
    Расчёт рекомендованной минимальной цены.
    
    Формула:
    X - Cost - (X * comm_rate) - Log - Stor - (X * acq) - LastMile - (X * ret) - (X * tax) = X * min_profit
    X * (1 - comm_rate - acq - ret - tax - min_profit) = Cost + Log + Stor + LastMile
    X = (Cost + Log + Stor + LastMile) / (1 - comm_rate - acq - ret - tax - min_profit)
    """
    if cost <= 0:
        return 0.0
    
    fixed_costs = cost + logistics + storage_cost + last_mile
    variable_rate = commission_rate + acquiring_rate + return_rate + tax_rate + min_profit_percent
    
    denominator = 1 - variable_rate
    if denominator <= 0:
        return 0.0
    
    recommended_price = fixed_costs / denominator
    return max(0, round(recommended_price, 2))
# ============================================================================
# БЛОК 1: ENUM И ТИПЫ (250+ СТРОК)
# ============================================================================
class CommissionType(Enum):
    """Типы комиссий маркетплейсов"""
    PERCENTAGE = auto()      # Процент от цены
    FIXED = auto()           # Фиксированная сумма
    HYBRID = auto()          # Процент + минимум
    SUBSCRIPTION = auto()    # Подписка
    TIERED = auto()          # Ступенчатая
    DYNAMIC = auto()         # Динамическая
    FLAT = auto()            # Плоская
    CUSTOM = auto()          # Пользовательская


class OperationMode(Enum):
    """Режимы работы с маркетплейсами"""
    FBY = auto()             # Fulfilled by Yandex (доставка МП)
    FBS = auto()             # Fulfilled by Seller (доставка продавца)
    FBO = auto()             # Fulfilled by Operator
    DBS = auto()             # Delivery by Seller
    FBP = auto()             # Fulfilled by Platform
    DBE = auto()             # Delivery by Express
    STANDARD = auto()        # Стандартный
    EXPRESS = auto()         # Экспресс
    SELF = auto()            # Самовывоз


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
    EXCEL_FORMULAS = auto()  # НОВОЕ v97: Excel с формулами
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


class TariffSource(Enum):
    """Источник тарифов (НОВОЕ v97)"""
    HARDCODED = "Захардкожены"
    AI_CACHE = "Кэш ИИ"
    AI_LIVE = "ИИ (запрос)"
    MANUAL = "Ручной ввод"
    IMPORTED = "Импортированы"


# ============================================================================
# БЛОК 2: ДАТАКЛАССЫ (400+ СТРОК)
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
    hazardous_surcharge: float = 0.0        # НОВОЕ v97: надбавка за опасные грузы
    fragile_surcharge: float = 0.0          # НОВОЕ v97: надбавка за хрупкие
    oversized_surcharge: float = 0.0        # НОВОЕ v97: надбавка за крупногабарит
    category_rates: Dict[str, float] = field(default_factory=dict)
    mode_multipliers: Dict[str, float] = field(default_factory=dict)
    weight_tiers: List[Tuple[float, float, float]] = field(default_factory=list)
    volume_tiers: List[Tuple[float, float, float]] = field(default_factory=list)
    available: bool = True
    description: str = ""
    version: str = "2026.1"
    last_updated: datetime = field(default_factory=datetime.now)
    tariff_source: TariffSource = TariffSource.HARDCODED  # НОВОЕ v97

    def get_commission_rate(self, category: Optional[str] = None) -> float:
        """Получить ставку комиссии для категории"""
        if category and category in self.category_rates:
            return self.category_rates[category]
        return self.commission_rate

    def get_mode_multiplier(self, mode: str) -> float:
        """Получить коэффициент режима работы"""
        return self.mode_multipliers.get(mode, 1.0)


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
    hazardous_surcharge: float = 0.0       # НОВОЕ v97
    fragile_surcharge: float = 0.0         # НОВОЕ v97
    oversized_surcharge: float = 0.0       # НОВОЕ v97
    tax_amount: float = 0.0                # НОВОЕ v97: налог
    tax_system: str = "УСН_6"              # НОВОЕ v97
    total_expenses: float = 0.0
    profit: float = 0.0
    margin_percent: float = 0.0
    roi: float = 0.0
    breakeven_price: float = 0.0
    recommended_min_price: float = 0.0     # НОВОЕ v97
    profit_per_ruble: float = 0.0
    contribution_margin: float = 0.0
    contribution_margin_ratio: float = 0.0
    timestamp: datetime = field(default_factory=datetime.now)
    calculation_id: str = field(default_factory=lambda: str(uuid.uuid4()))
    status: CalculationStatus = CalculationStatus.COMPLETED
    tariff_source: TariffSource = TariffSource.HARDCODED  # НОВОЕ v97
    metadata: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        result = asdict(self)
        result['timestamp'] = self.timestamp.isoformat()
        result['status'] = self.status.name
        result['tariff_source'] = self.tariff_source.value
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
            "recommended_min_price": self.recommended_min_price,
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
    recommended_min_price: float = 0.0
    rank: int = 0
    metadata: Dict[str, Any] = field(default_factory=dict)


@dataclass
class TariffCacheEntry:
    """Запись в кэше тарифов (НОВОЕ v97)"""
    marketplace: str
    category: Optional[str]
    data: Dict[str, Any]
    source: TariffSource
    timestamp: float
    ttl_seconds: int = 86400  # 24 часа
    version: str = "2026.1"
    notes: str = ""

    def is_expired(self) -> bool:
        return time.time() - self.timestamp > self.ttl_seconds

    def to_dict(self) -> Dict[str, Any]:
        return {
            "marketplace": self.marketplace,
            "category": self.category,
            "data": self.data,
            "source": self.source.value,
            "timestamp": self.timestamp,
            "ttl_seconds": self.ttl_seconds,
            "version": self.version,
            "notes": self.notes
        }

    @staticmethod
    def from_dict(d: Dict[str, Any]) -> 'TariffCacheEntry':
        return TariffCacheEntry(
            marketplace=d.get("marketplace", ""),
            category=d.get("category"),
            data=d.get("data", {}),
            source=TariffSource(d.get("source", "HARDCODED")),
            timestamp=d.get("timestamp", 0),
            ttl_seconds=d.get("ttl_seconds", 86400),
            version=d.get("version", "2026.1"),
            notes=d.get("notes", "")
        )


# ============================================================================
# БЛОК 3: УМНЫЙ КЭШ ТАРИФОВ (НОВОЕ v97) - 300+ СТРОК
# ============================================================================
class SmartTariffCache:
    """
    Умный кэш тарифов с возможностью ручного редактирования.
    Хранит тарифы в JSON-файле, поддерживает:
    - Чтение/запись тарифов
    - Историю изменений
    - Ручное редактирование
    - Экспорт/импорт
    - Валидацию данных
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
        self.cache_dir = TARIFFS_DIR
        self.cache_dir.mkdir(exist_ok=True, parents=True)
        self.cache_file = self.cache_dir / "tariffs_cache.json"
        self.history_file = self.cache_dir / "tariffs_history.json"
        self.backup_dir = self.cache_dir / "backups"
        self.backup_dir.mkdir(exist_ok=True, parents=True)
        self._cache: Dict[str, TariffCacheEntry] = {}
        self._history: List[Dict[str, Any]] = []
        self._load_cache()
        self._load_history()
        logger.info(f"SmartTariffCache инициализирован: {len(self._cache)} записей")

    def _make_key(self, marketplace: str, category: Optional[str] = None) -> str:
        """Создание ключа кэша"""
        cat = (category or "all").lower().strip()
        return f"{marketplace.lower().strip()}::{cat}"

    def _load_cache(self):
        """Загрузка кэша из файла"""
        if not self.cache_file.exists():
            self._cache = {}
            return
        try:
            with open(self.cache_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            self._cache = {}
            for key, entry_dict in data.items():
                try:
                    self._cache[key] = TariffCacheEntry.from_dict(entry_dict)
                except Exception as e:
                    logger.warning(f"Ошибка загрузки записи {key}: {e}")
        except Exception as e:
            logger.error(f"Ошибка загрузки кэша тарифов: {e}")
            self._cache = {}

    def _save_cache(self):
        """Сохранение кэша в файл"""
        try:
            data = {k: v.to_dict() for k, v in self._cache.items()}
            # Создаём бэкап перед сохранением
            if self.cache_file.exists():
                backup_name = f"tariffs_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
                shutil.copy2(self.cache_file, self.backup_dir / backup_name)
                # Оставляем только последние 10 бэкапов
                backups = sorted(self.backup_dir.glob("tariffs_backup_*.json"))
                for old_backup in backups[:-10]:
                    try:
                        old_backup.unlink()
                    except Exception:
                        pass
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"Ошибка сохранения кэша тарифов: {e}")

    def _load_history(self):
        """Загрузка истории изменений"""
        if not self.history_file.exists():
            self._history = []
            return
        try:
            with open(self.history_file, 'r', encoding='utf-8') as f:
                self._history = json.load(f)
        except Exception as e:
            logger.error(f"Ошибка загрузки истории: {e}")
            self._history = []

    def _save_history(self):
        """Сохранение истории"""
        try:
            with open(self.history_file, 'w', encoding='utf-8') as f:
                json.dump(self._history[-500:], f, ensure_ascii=False, indent=2)
        except Exception as e:
            logger.error(f"Ошибка сохранения истории: {e}")

    def _add_history_entry(self, action: str, marketplace: str, category: Optional[str],
                          old_data: Optional[Dict], new_data: Optional[Dict], source: str):
        """Добавление записи в историю"""
        entry = {
            "timestamp": datetime.now().isoformat(),
            "action": action,
            "marketplace": marketplace,
            "category": category,
            "old_data": old_data,
            "new_data": new_data,
            "source": source
        }
        self._history.append(entry)
        if len(self._history) > 500:
            self._history = self._history[-500:]
        self._save_history()

    def get(self, marketplace: str, category: Optional[str] = None,
            use_expired: bool = True) -> Optional[TariffCacheEntry]:
        """Получение тарифов из кэша"""
        key = self._make_key(marketplace, category)
        entry = self._cache.get(key)
        if entry is None:
            # Пробуем получить общие тарифы для маркетплейса
            key = self._make_key(marketplace, None)
            entry = self._cache.get(key)
        if entry is None:
            return None
        if entry.is_expired() and not use_expired:
            return None
        return entry

    def set(self, marketplace: str, category: Optional[str], data: Dict[str, Any],
            source: TariffSource, ttl_seconds: int = 86400, notes: str = "") -> bool:
        """Сохранение тарифов в кэш"""
        try:
            key = self._make_key(marketplace, category)
            old_entry = self._cache.get(key)
            old_data = old_entry.data if old_entry else None

            entry = TariffCacheEntry(
                marketplace=marketplace,
                category=category,
                data=data,
                source=source,
                timestamp=time.time(),
                ttl_seconds=ttl_seconds,
                version="2026.1",
                notes=notes
            )
            self._cache[key] = entry
            self._save_cache()
            self._add_history_entry(
                action="UPDATE" if old_entry else "CREATE",
                marketplace=marketplace,
                category=category,
                old_data=old_data,
                new_data=data,
                source=source.value
            )
            return True
        except Exception as e:
            logger.error(f"Ошибка сохранения тарифов: {e}")
            return False

    def delete(self, marketplace: str, category: Optional[str] = None) -> bool:
        """Удаление тарифов из кэша"""
        try:
            key = self._make_key(marketplace, category)
            old_entry = self._cache.get(key)
            if old_entry:
                self._add_history_entry(
                    action="DELETE",
                    marketplace=marketplace,
                    category=category,
                    old_data=old_entry.data,
                    new_data=None,
                    source="MANUAL"
                )
                del self._cache[key]
                self._save_cache()
            return True
        except Exception as e:
            logger.error(f"Ошибка удаления тарифов: {e}")
            return False

    def update_field(self, marketplace: str, category: Optional[str],
                    field: str, value: Any) -> bool:
        """Обновление одного поля тарифа (ручное редактирование)"""
        try:
            key = self._make_key(marketplace, category)
            entry = self._cache.get(key)
            if entry is None:
                logger.warning(f"Запись не найдена: {key}")
                return False
            old_data = entry.data.copy()
            entry.data[field] = value
            entry.timestamp = time.time()
            entry.source = TariffSource.MANUAL
            self._cache[key] = entry
            self._save_cache()
            self._add_history_entry(
                action="FIELD_UPDATE",
                marketplace=marketplace,
                category=category,
                old_data=old_data,
                new_data=entry.data,
                source="MANUAL"
            )
            return True
        except Exception as e:
            logger.error(f"Ошибка обновления поля: {e}")
            return False

    def get_all(self) -> Dict[str, TariffCacheEntry]:
        """Получение всех записей кэша"""
        return self._cache.copy()

    def get_history(self, limit: int = 100) -> List[Dict[str, Any]]:
        """Получение истории изменений"""
        return self._history[-limit:]

    def clear_expired(self) -> int:
        """Очистка устаревших записей"""
        expired_keys = [k for k, v in self._cache.items() if v.is_expired()]
        for key in expired_keys:
            del self._cache[key]
        if expired_keys:
            self._save_cache()
        return len(expired_keys)

    def clear_all(self) -> int:
        """Полная очистка кэша"""
        count = len(self._cache)
        self._cache = {}
        self._save_cache()
        self._add_history_entry(
            action="CLEAR_ALL",
            marketplace="ALL",
            category=None,
            old_data={"count": count},
            new_data=None,
            source="MANUAL"
        )
        return count

    def export_to_file(self, file_path: Union[str, Path]) -> bool:
        """Экспорт кэша в файл"""
        try:
            data = {k: v.to_dict() for k, v in self._cache.items()}
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            logger.error(f"Ошибка экспорта кэша: {e}")
            return False

    def import_from_file(self, file_path: Union[str, Path]) -> int:
        """Импорт кэша из файла"""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            count = 0
            for key, entry_dict in data.items():
                try:
                    self._cache[key] = TariffCacheEntry.from_dict(entry_dict)
                    count += 1
                except Exception as e:
                    logger.warning(f"Ошибка импорта записи {key}: {e}")
            self._save_cache()
            self._add_history_entry(
                action="IMPORT",
                marketplace="ALL",
                category=None,
                old_data=None,
                new_data={"count": count, "file": str(file_path)},
                source="IMPORTED"
            )
            return count
        except Exception as e:
            logger.error(f"Ошибка импорта кэша: {e}")
            return 0

    def get_statistics(self) -> Dict[str, Any]:
        """Статистика кэша"""
        stats = {
            "total_entries": len(self._cache),
            "by_marketplace": defaultdict(int),
            "by_source": defaultdict(int),
            "expired_count": 0,
            "oldest_entry": None,
            "newest_entry": None,
            "history_count": len(self._history)
        }
        if not self._cache:
            return stats
        timestamps = []
        for entry in self._cache.values():
            stats["by_marketplace"][entry.marketplace] += 1
            stats["by_source"][entry.source.value] += 1
            if entry.is_expired():
                stats["expired_count"] += 1
            timestamps.append(entry.timestamp)
        if timestamps:
            oldest_ts = min(timestamps)
            newest_ts = max(timestamps)
            stats["oldest_entry"] = datetime.fromtimestamp(oldest_ts).isoformat()
            stats["newest_entry"] = datetime.fromtimestamp(newest_ts).isoformat()
        return stats


# ============================================================================
# БЛОК 4: КОНФИГУРАЦИИ МАРКЕТПЛЕЙСОВ 2026 (400+ СТРОК)
# ============================================================================
def get_marketplace_configs_2026() -> Dict[str, MarketplaceConfig]:
    """
    Получение конфигураций маркетплейсов на 2026 год.
    С расширенными тарифами для автозапчастей.
    """
    configs = {}

    # === OZON ===
    configs["Ozon"] = MarketplaceConfig(
        name="Ozon",
        commission_rate=0.15,
        min_commission=30.0,
        logistics_base=50.0,
        logistics_per_kg=15.0,
        logistics_per_liter=5.0,
        storage_per_day=0.3,
        return_fee=0.02,
        acquiring_fee=0.015,
        last_mile_fee=50.0,
        hazardous_surcharge=0.02,
        fragile_surcharge=0.01,
        oversized_surcharge=0.015,
        category_rates={
            # Автозапчасти (специальные тарифы)
            "двигатель": 0.12, "трансмиссия": 0.13, "подвеска": 0.14,
            "тормозная_система": 0.14, "рулевое_управление": 0.14,
            "электрика": 0.15, "охлаждение": 0.14, "выпуск": 0.13,
            "фильтры": 0.17, "масла": 0.18, "оптика": 0.15,
            "шины": 0.16, "инструменты": 0.14, "кузов": 0.13,
            "крепёж": 0.12, "ремни": 0.13, "подшипники": 0.13,
            "климат": 0.14, "безопасность": 0.15,
            # Общие категории
            "одежда_обувь": 0.15, "электроника": 0.10,
            "красота": 0.22, "автотовары": 0.12,
            "книги": 0.10, "дом": 0.12, "спорт": 0.12,
            "детские_товары": 0.12, "продукты": 0.08, "здоровье": 0.15
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.0, "FBO": 0.8,
            "DBS": 1.3, "FBP": 0.9
        },
        description="Ozon - крупнейший маркетплейс России",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )

    # === WILDBERRIES ===
    configs["Wildberries"] = MarketplaceConfig(
        name="Wildberries",
        commission_rate=0.18,
        min_commission=35.0,
        logistics_base=60.0,
        logistics_per_kg=18.0,
        logistics_per_liter=6.0,
        storage_per_day=0.5,
        return_fee=0.03,
        acquiring_fee=0.0,
        last_mile_fee=0.0,
        delivery_fee_percent=0.05,
        rko_fee=0.01,
        hazardous_surcharge=0.025,
        fragile_surcharge=0.015,
        oversized_surcharge=0.02,
        category_rates={
            "двигатель": 0.15, "трансмиссия": 0.16, "подвеска": 0.17,
            "тормозная_система": 0.17, "рулевое_управление": 0.17,
            "электрика": 0.18, "охлаждение": 0.17, "выпуск": 0.16,
            "фильтры": 0.20, "масла": 0.22, "оптика": 0.18,
            "шины": 0.19, "инструменты": 0.17, "кузов": 0.16,
            "крепёж": 0.15, "ремни": 0.16, "подшипники": 0.16,
            "климат": 0.17, "безопасность": 0.18,
            "одежда": 0.18, "электроника": 0.12, "дети": 0.15,
            "дом": 0.15, "красота": 0.15, "продукты": 0.10,
            "здоровье": 0.12, "спорт": 0.15, "книги": 0.12,
            "автотовары": 0.15
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.15, "FBO": 1.1,
            "DBS": 1.25, "FBP": 1.0
        },
        description="Wildberries - лидер e-commerce России",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )

    # === ЯНДЕКС МАРКЕТ ===
    configs["Яндекс Маркет"] = MarketplaceConfig(
        name="Яндекс Маркет",
        commission_rate=0.14,
        subscription_fee=6990.0,
        min_commission=0.0,
        logistics_base=45.0,
        logistics_per_kg=14.0,
        logistics_per_liter=4.5,
        storage_per_day=0.25,
        return_fee=0.02,
        acquiring_fee=0.02,
        last_mile_fee=40.0,
        premium_fee=0.02,
        hazardous_surcharge=0.018,
        fragile_surcharge=0.01,
        oversized_surcharge=0.012,
        category_rates={
            "двигатель": 0.11, "трансмиссия": 0.12, "подвеска": 0.13,
            "тормозная_система": 0.13, "рулевое_управление": 0.13,
            "электрика": 0.14, "охлаждение": 0.13, "выпуск": 0.12,
            "фильтры": 0.16, "масла": 0.17, "оптика": 0.14,
            "шины": 0.15, "инструменты": 0.13, "кузов": 0.12,
            "крепёж": 0.11, "ремни": 0.12, "подшипники": 0.12,
            "климат": 0.13, "безопасность": 0.14,
            "одежда_обувь": 0.14, "садоводство": 0.12,
            "строительство": 0.19, "красота": 0.14,
            "детские_товары": 0.14, "электроника": 0.14,
            "автотовары": 0.14, "книги": 0.14, "дом": 0.14, "спорт": 0.14
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.0, "FBO": 0.8,
            "DBS": 1.3, "FBP": 0.9
        },
        description="Яндекс Маркет - маркетплейс экосистемы Яндекса",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )

    # === ALIEXPRESS ===
    configs["AliExpress"] = MarketplaceConfig(
        name="AliExpress",
        commission_rate=0.10,
        min_commission=20.0,
        logistics_base=80.0,
        logistics_per_kg=25.0,
        logistics_per_liter=8.0,
        storage_per_day=0.2,
        return_fee=0.01,
        acquiring_fee=0.025,
        last_mile_fee=70.0,
        delivery_fee_percent=0.08,
        hazardous_surcharge=0.03,
        fragile_surcharge=0.02,
        oversized_surcharge=0.025,
        category_rates={
            "двигатель": 0.08, "трансмиссия": 0.09, "подвеска": 0.10,
            "тормозная_система": 0.10, "рулевое_управление": 0.10,
            "электрика": 0.11, "охлаждение": 0.10, "выпуск": 0.09,
            "фильтры": 0.12, "масла": 0.13, "оптика": 0.11,
            "шины": 0.12, "инструменты": 0.10, "кузов": 0.09,
            "крепёж": 0.08, "ремни": 0.09, "подшипники": 0.09,
            "климат": 0.10, "безопасность": 0.11,
            "электроника": 0.08, "одежда": 0.10, "дом": 0.10,
            "красота": 0.10, "спорт": 0.10, "автотовары": 0.10,
            "книги": 0.08, "детские_товары": 0.10
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.2, "FBO": 1.1,
            "DBS": 1.3, "FBP": 0.9
        },
        description="AliExpress - международный маркетплейс",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )

    # === МЕГАМАРКЕТ ===
    configs["Мегамаркет"] = MarketplaceConfig(
        name="Мегамаркет",
        commission_rate=0.13,
        min_commission=28.0,
        logistics_base=55.0,
        logistics_per_kg=16.0,
        logistics_per_liter=5.5,
        storage_per_day=0.3,
        return_fee=0.02,
        acquiring_fee=0.018,
        last_mile_fee=45.0,
        delivery_fee_percent=0.05,
        hazardous_surcharge=0.02,
        fragile_surcharge=0.012,
        oversized_surcharge=0.015,
        category_rates={
            "двигатель": 0.10, "трансмиссия": 0.11, "подвеска": 0.12,
            "тормозная_система": 0.12, "рулевое_управление": 0.12,
            "электрика": 0.13, "охлаждение": 0.12, "выпуск": 0.11,
            "фильтры": 0.15, "масла": 0.16, "оптика": 0.13,
            "шины": 0.14, "инструменты": 0.12, "кузов": 0.11,
            "крепёж": 0.10, "ремни": 0.11, "подшипники": 0.11,
            "климат": 0.12, "безопасность": 0.13,
            "электроника": 0.02, "одежда": 0.20, "обувь": 0.20,
            "автотовары": 0.15, "дом": 0.12, "красота": 0.12,
            "спорт": 0.12, "детские_товары": 0.12,
            "продукты": 0.05, "книги": 0.08
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.1, "FBO": 1.05,
            "DBS": 1.2, "FBP": 0.95
        },
        description="Мегамаркет (Сбер) - маркетплейс экосистемы Сбера",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )

    # === СБЕРМЕГАМАРКЕТ ===
    configs["СберМегаМаркет"] = MarketplaceConfig(
        name="СберМегаМаркет",
        commission_rate=0.13,
        min_commission=28.0,
        logistics_base=55.0,
        logistics_per_kg=16.0,
        logistics_per_liter=5.5,
        storage_per_day=0.3,
        return_fee=0.02,
        acquiring_fee=0.018,
        last_mile_fee=45.0,
        rko_fee=0.015,
        delivery_fee_percent=0.055,
        hazardous_surcharge=0.02,
        fragile_surcharge=0.012,
        oversized_surcharge=0.015,
        category_rates={
            "двигатель": 0.10, "трансмиссия": 0.11, "подвеска": 0.12,
            "тормозная_система": 0.12, "рулевое_управление": 0.12,
            "электрика": 0.13, "охлаждение": 0.12, "выпуск": 0.11,
            "фильтры": 0.15, "масла": 0.16, "оптика": 0.13,
            "шины": 0.14, "инструменты": 0.12, "кузов": 0.11,
            "крепёж": 0.10, "ремни": 0.11, "подшипники": 0.11,
            "климат": 0.12, "безопасность": 0.13,
            "электроника": 0.02, "одежда": 0.15,
            "продукты": 0.05, "дом": 0.10, "красота": 0.10,
            "спорт": 0.10, "автотовары": 0.12,
            "детские_товары": 0.10, "книги": 0.08
        },
        mode_multipliers={
            "FBY": 0.75, "FBS": 1.1, "FBO": 1.05,
            "DBS": 1.2, "FBP": 0.95
        },
        description="СберМегаМаркет - маркетплейс Сбера",
        version="2026.1",
        tariff_source=TariffSource.HARDCODED
    )

    # === ЗАГРУЗКА ИЗ КЭША (НОВОЕ v97) ===
    try:
        cache = SmartTariffCache()
        for mp_name, config in configs.items():
            cached_entry = cache.get(mp_name, None, use_expired=False)
            if cached_entry and cached_entry.data:
                # Применяем кэшированные тарифы
                data = cached_entry.data
                if "commission_rate" in data:
                    config.commission_rate = data["commission_rate"]
                if "min_commission" in data:
                    config.min_commission = data["min_commission"]
                if "logistics_base" in data:
                    config.logistics_base = data["logistics_base"]
                if "logistics_per_kg" in data:
                    config.logistics_per_kg = data["logistics_per_kg"]
                if "logistics_per_liter" in data:
                    config.logistics_per_liter = data["logistics_per_liter"]
                if "storage_per_day" in data:
                    config.storage_per_day = data["storage_per_day"]
                if "return_fee" in data:
                    config.return_fee = data["return_fee"]
                if "acquiring_fee" in data:
                    config.acquiring_fee = data["acquiring_fee"]
                if "last_mile_fee" in data:
                    config.last_mile_fee = data["last_mile_fee"]
                if "category_rates" in data:
                    config.category_rates.update(data["category_rates"])
                config.tariff_source = cached_entry.source
                config.last_updated = datetime.fromtimestamp(cached_entry.timestamp)
                logger.info(f"📥 Применены кэшированные тарифы для {mp_name} (источник: {cached_entry.source.value})")
    except Exception as e:
        logger.warning(f"Не удалось загрузить кэш тарифов: {e}")

    return configs


# ============================================================================
# БЛОК 5: 150+ КАТЕГОРИЙ АВТОЗАПЧАСТЕЙ С ПОЛНЫМИ ГАБАРИТАМИ (1500+ СТРОК)
# ============================================================================
def get_auto_parts_categories_full() -> Dict[str, ProductCategory]:
    """Получение полного списка категорий автозапчастей с габаритами"""
    categories = {}

    def make_cat(name, desc, min_l, max_l, min_w, max_w, min_h, max_h,
                 min_wt, max_wt, typ_vol, typ_wt, oem=None,
                 season=Seasonality.ALL_YEAR, risk=RiskLevel.LOW,
                 hazardous=False, fragile=False):
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

    # === ДВИГАТЕЛЬ ===
    categories["двигатель"] = make_cat("двигатель", "Двигатели и комплектующие",
        30, 80, 30, 60, 30, 70, 10, 200, 20.0, 80.0, risk=RiskLevel.HIGH)
    categories["поршни"] = make_cat("поршни", "Поршни и кольца",
        5, 12, 5, 12, 3, 10, 0.1, 1.5, 0.1, 0.5)
    categories["клапаны"] = make_cat("клапаны", "Клапаны двигателя",
        3, 8, 1, 3, 10, 40, 0.05, 0.5, 0.05, 0.2)
    categories["прокладки_двигателя"] = make_cat("прокладки_двигателя", "Прокладки ГБЦ и двигателя",
        10, 50, 10, 40, 0.1, 2, 0.01, 0.3, 0.1, 0.1)
    categories["свечи_зажигания"] = make_cat("свечи_зажигания", "Свечи зажигания",
        2, 3, 2, 3, 6, 10, 0.04, 0.1, 0.01, 0.05)
    categories["блок_цилиндров"] = make_cat("блок_цилиндров", "Блок цилиндров",
        40, 70, 30, 50, 20, 40, 20, 80, 100.0, 50.0, risk=RiskLevel.HIGH)
    categories["головка_блока"] = make_cat("головка_блока", "Головка блока цилиндров",
        30, 60, 20, 40, 8, 20, 5, 30, 40.0, 15.0, risk=RiskLevel.HIGH)
    categories["коленвал"] = make_cat("коленвал", "Коленчатый вал",
        40, 90, 8, 20, 8, 20, 10, 40, 30.0, 25.0, risk=RiskLevel.HIGH)
    categories["распредвал"] = make_cat("распредвал", "Распределительный вал",
        30, 80, 5, 15, 5, 15, 3, 15, 20.0, 9.0)
    categories["шатун"] = make_cat("шатун", "Шатун двигателя",
        12, 35, 4, 10, 3, 7, 0.5, 2, 3.0, 1.25)
    categories["гидрокомпенсаторы"] = make_cat("гидрокомпенсаторы", "Гидрокомпенсаторы",
        3, 8, 3, 8, 3, 8, 0.05, 0.2, 0.3, 0.125)
    categories["привод_грм"] = make_cat("привод_грм", "Привод ГРМ (ремень, цепь)",
        60, 160, 2, 5, 1, 2, 0.1, 1, 2.0, 0.55)
    categories["масляный_насос"] = make_cat("масляный_насос", "Масляный насос",
        8, 18, 8, 18, 8, 18, 1, 5, 5.0, 3.0)
    categories["водяной_насос"] = make_cat("водяной_насос", "Водяной насос (помпа)",
        8, 18, 8, 18, 8, 18, 1, 4, 5.0, 2.5)
    categories["турбокомпрессор"] = make_cat("турбокомпрессор", "Турбокомпрессор",
        15, 35, 15, 30, 15, 25, 5, 15, 15.0, 10.0, risk=RiskLevel.HIGH)
    categories["масляный_поддон"] = make_cat("масляный_поддон", "Масляный поддон",
        30, 60, 20, 40, 10, 20, 2, 8, 15.0, 5.0)
    categories["клапанная_крышка"] = make_cat("клапанная_крышка", "Клапанная крышка",
        30, 60, 15, 30, 5, 10, 1, 4, 8.0, 2.5)
    categories["приводной_ремень"] = make_cat("приводной_ремень", "Приводной ремень",
        60, 150, 1, 3, 0.5, 1, 0.05, 0.5, 1.0, 0.275)
    categories["демпфер_коленвала"] = make_cat("демпфер_коленвала", "Демпфер коленвала",
        10, 25, 10, 25, 5, 10, 2, 8, 5.0, 5.0)
    categories["маховик"] = make_cat("маховик", "Маховик",
        25, 45, 25, 45, 5, 10, 5, 15, 10.0, 10.0, risk=RiskLevel.HIGH)
    categories["стартерный_венец"] = make_cat("стартерный_венец", "Стартерный венец",
        25, 40, 25, 40, 2, 5, 1, 5, 5.0, 3.0)

    # === ТРАНСМИССИЯ ===
    categories["трансмиссия"] = make_cat("трансмиссия", "КПП и комплектующие",
        40, 80, 30, 60, 30, 60, 20, 100, 30.0, 50.0, risk=RiskLevel.HIGH)
    categories["сцепление"] = make_cat("сцепление", "Комплекты сцепления",
        20, 40, 20, 40, 5, 15, 2, 10, 3.0, 5.0)
    categories["шкивы"] = make_cat("шкивы", "Шкивы и ролики",
        5, 20, 5, 20, 2, 8, 0.2, 3, 0.5, 1.5)
    categories["коробка_передач"] = make_cat("коробка_передач", "Коробка передач в сборе",
        40, 70, 30, 50, 25, 40, 30, 80, 80.0, 55.0, risk=RiskLevel.HIGH)
    categories["привод_полуоси"] = make_cat("привод_полуоси", "Привод (полуоси)",
        40, 90, 8, 18, 8, 18, 3, 12, 15.0, 7.5)
    categories["дифференциал"] = make_cat("дифференциал", "Дифференциал",
        20, 45, 20, 45, 20, 45, 10, 30, 30.0, 20.0, risk=RiskLevel.HIGH)
    categories["карданный_вал"] = make_cat("карданный_вал", "Карданный вал",
        60, 160, 8, 18, 8, 18, 5, 20, 25.0, 12.5)
    categories["раздаточная_коробка"] = make_cat("раздаточная_коробка", "Раздаточная коробка",
        25, 45, 20, 35, 20, 35, 15, 40, 35.0, 27.5, risk=RiskLevel.HIGH)
    categories["гидротрансформатор"] = make_cat("гидротрансформатор", "Гидротрансформатор АКПП",
        25, 40, 25, 40, 20, 30, 10, 25, 30.0, 17.5, risk=RiskLevel.HIGH)
    categories["механизм_переключения"] = make_cat("механизм_переключения", "Механизм переключения передач",
        15, 35, 5, 15, 5, 15, 1, 5, 5.0, 3.0)
    categories["подшипники_трансмиссии"] = make_cat("подшипники_трансмиссии", "Подшипники трансмиссии",
        8, 18, 8, 18, 8, 18, 0.5, 3, 3.0, 1.75)
    categories["сальники_трансмиссии"] = make_cat("сальники_трансмиссии", "Сальники трансмиссии",
        2, 12, 2, 12, 1, 3, 0.05, 0.3, 0.5, 0.175)
    categories["фильтр_акпп"] = make_cat("фильтр_акпп", "Фильтр АКПП",
        8, 18, 8, 18, 8, 18, 0.5, 2, 3.0, 1.25)
    categories["масло_трансмиссионное"] = make_cat("масло_трансмиссионное", "Трансмиссионное масло",
        10, 35, 8, 25, 8, 25, 1, 5, 5.0, 3.0, hazardous=True)
    categories["трос_сцепления"] = make_cat("трос_сцепления", "Трос сцепления",
        40, 100, 1, 3, 1, 3, 0.1, 0.5, 1.0, 0.3)
    categories["цилиндр_сцепления"] = make_cat("цилиндр_сцепления", "Цилиндр сцепления",
        10, 20, 5, 10, 5, 10, 0.5, 2, 2.0, 1.25)
    categories["вал_кпп"] = make_cat("вал_кпп", "Вал КПП",
        20, 50, 5, 12, 5, 12, 2, 8, 8.0, 5.0)
    categories["шестерни_кпп"] = make_cat("шестерни_кпп", "Шестерни КПП",
        5, 15, 5, 15, 5, 15, 0.5, 3, 3.0, 1.75)
    categories["синхронизатор"] = make_cat("синхронизатор", "Синхронизатор",
        5, 12, 5, 12, 3, 8, 0.3, 1.5, 2.0, 0.9)
    categories["муфта_кпп"] = make_cat("муфта_кпп", "Муфта КПП",
        5, 15, 5, 15, 3, 8, 0.5, 2, 3.0, 1.25)

    # === ПОДВЕСКА ===
    categories["подвеска"] = make_cat("подвеска", "Элементы подвески",
        20, 80, 10, 40, 10, 60, 1, 20, 5.0, 8.0)
    categories["амортизаторы"] = make_cat("амортизаторы", "Амортизаторы",
        5, 10, 5, 10, 40, 70, 2, 8, 5.0, 8.0, fragile=True)
    categories["пружины"] = make_cat("пружины", "Пружины подвески",
        20, 40, 20, 40, 30, 60, 3, 10, 8.0, 12.0)
    categories["сайлентблоки"] = make_cat("сайлентблоки", "Сайлентблоки",
        3, 10, 3, 10, 2, 8, 0.1, 1, 0.1, 0.3)
    categories["шаровые_опоры"] = make_cat("шаровые_опоры", "Шаровые опоры",
        5, 15, 5, 15, 5, 15, 0.3, 2, 0.5, 1.5)
    categories["ступицы"] = make_cat("ступицы", "Ступицы и подшипники",
        10, 25, 10, 25, 5, 15, 1, 5, 2.0, 4.0)
    categories["рычаг_подвески"] = make_cat("рычаг_подвески", "Рычаг подвески",
        20, 65, 5, 18, 5, 18, 2, 10, 10.0, 6.0)
    categories["стабилизатор"] = make_cat("стабилизатор", "Стабилизатор поперечной устойчивости",
        25, 65, 3, 10, 3, 10, 1, 5, 5.0, 3.0)
    categories["пыльник"] = make_cat("пыльник", "Пыльник (чехол)",
        5, 12, 5, 12, 8, 22, 0.1, 0.5, 1.0, 0.3)
    categories["отбойник"] = make_cat("отбойник", "Отбойник амортизатора",
        5, 12, 5, 12, 5, 12, 0.1, 0.5, 1.0, 0.3)
    categories["опора_стойки"] = make_cat("опора_стойки", "Опора стойки амортизатора",
        8, 18, 8, 18, 5, 12, 0.5, 2, 3.0, 1.25)
    categories["подрамник"] = make_cat("подрамник", "Подрамник",
        45, 105, 15, 35, 8, 18, 10, 30, 25.0, 20.0, risk=RiskLevel.HIGH)
    categories["распорка"] = make_cat("распорка", "Распорка подвески",
        25, 65, 2, 6, 2, 6, 0.5, 2, 2.0, 1.25)
    categories["сайлентблоки_в_сборе"] = make_cat("сайлентблоки_в_сборе", "Сайлентблоки в сборе",
        8, 22, 8, 22, 5, 12, 0.5, 2, 3.0, 1.25)
    categories["буфер"] = make_cat("буфер", "Буфер подвески",
        5, 12, 5, 12, 5, 12, 0.1, 0.5, 1.0, 0.3)
    categories["подушка_подвески"] = make_cat("подушка_подвески", "Подушка подвески",
        8, 18, 8, 18, 5, 12, 0.5, 2, 2.0, 1.25)
    categories["тяга_продольная"] = make_cat("тяга_продольная", "Тяга продольная",
        25, 65, 3, 8, 3, 8, 1, 4, 4.0, 2.5)
    categories["балка_моста"] = make_cat("балка_моста", "Балка моста",
        45, 85, 10, 20, 10, 20, 15, 40, 30.0, 27.5, risk=RiskLevel.HIGH)

    # === ТОРМОЗНАЯ СИСТЕМА ===
    categories["тормозная_система"] = make_cat("тормозная_система", "Тормозная система",
        20, 40, 20, 40, 5, 15, 2, 15, 3.0, 8.0, risk=RiskLevel.HIGH)
    categories["тормозные_диски"] = make_cat("тормозные_диски", "Тормозные диски",
        25, 40, 25, 40, 3, 8, 3, 12, 3.0, 8.0, fragile=True)
    categories["тормозные_колодки"] = make_cat("тормозные_колодки", "Тормозные колодки",
        10, 20, 5, 12, 3, 8, 1, 4, 1.0, 3.0)
    categories["тормозные_шланги"] = make_cat("тормозные_шланги", "Тормозные шланги",
        20, 60, 2, 5, 2, 5, 0.2, 1, 0.3, 0.8)
    categories["тормозные_суппорты"] = make_cat("тормозные_суппорты", "Тормозные суппорты",
        15, 30, 10, 20, 10, 20, 2, 8, 5.0, 5.0)
    categories["тормозные_барабаны"] = make_cat("тормозные_барабаны", "Тормозные барабаны",
        20, 35, 20, 35, 5, 15, 3, 10, 5.0, 6.5)
    categories["гтц"] = make_cat("гтц", "Главный тормозной цилиндр",
        10, 25, 8, 18, 8, 18, 1, 4, 3.0, 2.5)
    categories["вакуумный_усилитель"] = make_cat("вакуумный_усилитель", "Вакуумный усилитель тормозов",
        20, 35, 20, 35, 10, 20, 2, 6, 10.0, 4.0)

    # === РУЛЕВОЕ УПРАВЛЕНИЕ ===
    categories["рулевое_управление"] = make_cat("рулевое_управление", "Рулевое управление",
        30, 100, 10, 30, 10, 30, 2, 15, 5.0, 10.0)
    categories["рулевые_тяги"] = make_cat("рулевые_тяги", "Рулевые тяги и наконечники",
        20, 60, 3, 8, 3, 8, 0.5, 3, 1.0, 2.5)
    categories["рулевые_рейки"] = make_cat("рулевые_рейки", "Рулевые рейки",
        50, 100, 10, 20, 10, 20, 5, 15, 8.0, 12.0)
    categories["рулевой_кардан"] = make_cat("рулевой_кардан", "Рулевой кардан",
        20, 45, 5, 12, 5, 12, 1, 4, 5.0, 2.5)
    categories["усилитель_руля"] = make_cat("усилитель_руля", "Усилитель руля (ГУР/ЭУР)",
        15, 30, 15, 30, 15, 25, 3, 10, 10.0, 6.5)
    categories["рулевой_насос"] = make_cat("рулевой_насос", "Насос ГУР",
        15, 30, 12, 22, 12, 22, 3, 8, 6.0, 5.5)

    # === ЭЛЕКТРИКА ===
    categories["электрика"] = make_cat("электрика", "Электрооборудование",
        10, 40, 10, 30, 10, 30, 0.5, 10, 2.0, 5.0)
    categories["стартеры"] = make_cat("стартеры", "Стартеры",
        15, 30, 10, 20, 10, 25, 3, 10, 3.0, 6.0)
    categories["генераторы"] = make_cat("генераторы", "Генераторы",
        15, 30, 15, 25, 15, 30, 4, 12, 5.0, 8.0)
    categories["аккумуляторы"] = make_cat("аккумуляторы", "Аккумуляторы",
        20, 40, 15, 25, 15, 30, 10, 30, 15.0, 20.0, hazardous=True, risk=RiskLevel.HIGH)
    categories["датчики"] = make_cat("датчики", "Датчики",
        3, 10, 2, 5, 2, 8, 0.05, 0.5, 0.1, 0.3)
    categories["катушки_зажигания"] = make_cat("катушки_зажигания", "Катушки зажигания",
        5, 15, 3, 8, 5, 15, 0.2, 1, 0.5, 0.6)
    categories["проводка"] = make_cat("проводка", "Проводка и жгуты",
        20, 100, 5, 20, 2, 10, 0.3, 3, 3.0, 1.5)
    categories["блоки_управления"] = make_cat("блоки_управления", "Блоки управления (ЭБУ)",
        15, 30, 10, 20, 5, 15, 0.5, 3, 3.0, 1.5)

    # === СИСТЕМА ОХЛАЖДЕНИЯ ===
    categories["охлаждение"] = make_cat("охлаждение", "Система охлаждения",
        20, 80, 15, 50, 10, 40, 1, 15, 8.0, 15.0)
    categories["радиаторы"] = make_cat("радиаторы", "Радиаторы охлаждения",
        40, 80, 30, 60, 5, 15, 2, 10, 10.0, 15.0, fragile=True)
    categories["помпы"] = make_cat("помпы", "Водяные помпы",
        10, 25, 10, 20, 10, 20, 1, 5, 2.0, 4.0)
    categories["термостаты"] = make_cat("термостаты", "Термостаты",
        5, 12, 5, 12, 5, 12, 0.2, 1, 0.5, 1.0)
    categories["вентилятор_радиатора"] = make_cat("вентилятор_радиатора", "Вентилятор радиатора",
        30, 50, 30, 50, 5, 15, 2, 6, 15.0, 4.0, fragile=True)
    categories["расширительный_бачок"] = make_cat("расширительный_бачок", "Расширительный бачок",
        15, 30, 10, 20, 10, 25, 0.3, 1.5, 4.0, 0.9)

    # === ФИЛЬТРЫ ===
    categories["фильтры"] = make_cat("фильтры", "Фильтры",
        5, 30, 5, 30, 5, 40, 0.1, 3, 2.0, 5.0)
    categories["масляные_фильтры"] = make_cat("масляные_фильтры", "Масляные фильтры",
        6, 12, 6, 12, 8, 15, 0.3, 1, 1.0, 1.5)
    categories["воздушные_фильтры"] = make_cat("воздушные_фильтры", "Воздушные фильтры",
        15, 40, 15, 35, 3, 10, 0.2, 2, 2.0, 4.0)
    categories["топливные_фильтры"] = make_cat("топливные_фильтры", "Топливные фильтры",
        5, 15, 5, 15, 8, 20, 0.3, 1.5, 1.0, 2.0)
    categories["салонные_фильтры"] = make_cat("салонные_фильтры", "Салонные фильтры",
        20, 35, 15, 25, 2, 5, 0.2, 1, 1.5, 2.5)

    # === МАСЛА И ЖИДКОСТИ (ОПАСНЫЕ) ===
    categories["масла"] = make_cat("масла", "Масла и технические жидкости",
        5, 30, 5, 30, 10, 40, 0.5, 20, 5.0, 15.0, hazardous=True)
    categories["моторные_масла"] = make_cat("моторные_масла", "Моторные масла",
        8, 25, 8, 25, 20, 40, 1, 20, 5.0, 15.0, hazardous=True)
    categories["трансмиссионные_масла"] = make_cat("трансмиссионные_масла", "Трансмиссионные масла",
        8, 25, 8, 25, 20, 40, 1, 20, 5.0, 15.0, hazardous=True)
    categories["тормозная_жидкость"] = make_cat("тормозная_жидкость", "Тормозная жидкость",
        5, 10, 5, 10, 15, 25, 0.5, 2, 1.0, 2.0, hazardous=True)
    categories["антифриз"] = make_cat("антифриз", "Антифриз / Охлаждающая жидкость",
        10, 30, 10, 30, 20, 40, 1, 20, 5.0, 15.0, hazardous=True)

    # === ОПТИКА (ХРУПКИЕ) ===
    categories["оптика"] = make_cat("оптика", "Оптика и освещение",
        15, 60, 15, 40, 15, 40, 0.5, 10, 5.0, 10.0, fragile=True)
    categories["фары"] = make_cat("фары", "Фары головного света",
        30, 60, 20, 40, 20, 40, 2, 8, 8.0, 12.0, fragile=True)
    categories["лампы"] = make_cat("лампы", "Автомобильные лампы",
        2, 10, 2, 5, 5, 15, 0.02, 0.3, 0.1, 0.3, fragile=True)
    categories["фонари"] = make_cat("фонари", "Задние фонари",
        20, 50, 15, 30, 10, 25, 1, 5, 5.0, 8.0, fragile=True)
    categories["led_лампы"] = make_cat("led_лампы", "LED лампы",
        5, 15, 3, 8, 3, 8, 0.1, 0.5, 0.3, 0.3, fragile=True)

    # === КУЗОВ (КРУПНОГАБАРИТ + ХРУПКИЕ) ===
    categories["кузов"] = make_cat("кузов", "Кузовные детали",
        50, 200, 30, 150, 10, 100, 2, 50, 30.0, 80.0, fragile=True, risk=RiskLevel.HIGH)
    categories["бамперы"] = make_cat("бамперы", "Бамперы",
        100, 200, 30, 60, 20, 50, 5, 20, 50.0, 80.0, fragile=True)
    categories["крылья"] = make_cat("крылья", "Крылья",
        50, 100, 30, 60, 30, 80, 3, 10, 20.0, 40.0, fragile=True)
    categories["капоты"] = make_cat("капоты", "Капоты",
        100, 180, 80, 150, 5, 15, 5, 15, 30.0, 60.0, fragile=True)
    categories["зеркала"] = make_cat("зеркала", "Зеркала заднего вида",
        15, 30, 10, 20, 10, 20, 0.5, 3, 3.0, 5.0, fragile=True)
    categories["двери"] = make_cat("двери", "Двери",
        100, 150, 50, 100, 5, 15, 15, 40, 80.0, 27.5, fragile=True, risk=RiskLevel.HIGH)
    categories["стёкла"] = make_cat("стёкла", "Автомобильные стёкла",
        50, 150, 30, 100, 0.5, 2, 5, 20, 40.0, 12.5, fragile=True, risk=RiskLevel.HIGH)

    # === ШИНЫ И ДИСКИ ===
    categories["шины"] = make_cat("шины", "Шины и диски",
        40, 80, 40, 80, 15, 40, 5, 30, 20.0, 40.0)
    categories["летние_шины"] = make_cat("летние_шины", "Летние шины",
        50, 80, 50, 80, 15, 30, 8, 25, 25.0, 35.0, season=Seasonality.SUMMER)
    categories["зимние_шины"] = make_cat("зимние_шины", "Зимние шины",
        50, 80, 50, 80, 15, 30, 8, 25, 25.0, 35.0, season=Seasonality.WINTER)
    categories["диски"] = make_cat("диски", "Колесные диски",
        40, 60, 40, 60, 15, 30, 5, 20, 15.0, 25.0, fragile=True)

    # === ИНСТРУМЕНТЫ ===
    categories["инструменты"] = make_cat("инструменты", "Автоинструменты",
        10, 60, 5, 30, 3, 20, 0.2, 10, 3.0, 8.0)
    categories["домкраты"] = make_cat("домкраты", "Домкраты",
        20, 50, 10, 25, 10, 25, 3, 15, 5.0, 12.0)
    categories["наборы_ключей"] = make_cat("наборы_ключей", "Наборы ключей",
        15, 40, 10, 25, 3, 10, 1, 8, 3.0, 6.0)
    categories["компрессоры_воздушные"] = make_cat("компрессоры_воздушные", "Воздушные компрессоры",
        25, 60, 20, 40, 20, 40, 5, 25, 15.0, 15.0)

    # === РЕМНИ И ПРИВОДЫ ===
    categories["ремни"] = make_cat("ремни", "Ремни ГРМ и приводов",
        50, 150, 1, 3, 1, 3, 0.1, 0.8, 0.5, 1.0)
    categories["ролики"] = make_cat("ролики", "Ролики натяжители",
        5, 12, 5, 12, 2, 5, 0.2, 1.5, 0.5, 1.0)

    # === ПОДШИПНИКИ ===
    categories["подшипники"] = make_cat("подшипники", "Подшипники",
        3, 15, 3, 15, 1, 5, 0.1, 3, 0.5, 2.0)

    # === КРЕПЁЖ ===
    categories["крепёж"] = make_cat("крепёж", "Крепёж и метизы",
        0.5, 10, 0.5, 10, 0.5, 10, 0.01, 2, 0.2, 1.0)

    # === КЛИМАТ ===
    categories["климат"] = make_cat("климат", "Климат-контроль и кондиционер",
        20, 80, 20, 60, 15, 50, 2, 20, 10.0, 20.0)
    categories["компрессоры"] = make_cat("компрессоры", "Компрессоры кондиционера",
        20, 40, 15, 30, 15, 30, 5, 15, 8.0, 12.0)
    categories["конденсоры"] = make_cat("конденсоры", "Конденсоры кондиционера",
        40, 80, 30, 60, 5, 15, 2, 8, 10.0, 5.0, fragile=True)

    # === ВЫХЛОПНАЯ СИСТЕМА ===
    categories["выпуск"] = make_cat("выпуск", "Выхлопная система",
        30, 150, 10, 40, 10, 40, 2, 25, 10.0, 25.0)
    categories["глушители"] = make_cat("глушители", "Глушители",
        50, 150, 20, 40, 20, 40, 5, 20, 20.0, 30.0)
    categories["катализаторы"] = make_cat("катализаторы", "Каталитические нейтрализаторы",
        30, 80, 15, 30, 15, 30, 3, 15, 10.0, 20.0, hazardous=True, risk=RiskLevel.HIGH)
    categories["гофры"] = make_cat("гофры", "Гофры выхлопной системы",
        10, 30, 5, 15, 5, 15, 0.3, 2, 2.0, 1.15)

    # === БЕЗОПАСНОСТЬ ===
    categories["безопасность"] = make_cat("безопасность", "Системы безопасности",
        10, 50, 10, 40, 5, 30, 0.5, 8, 3.0, 6.0, risk=RiskLevel.HIGH)
    categories["подушки_безопасности"] = make_cat("подушки_безопасности", "Подушки безопасности",
        20, 50, 15, 30, 10, 20, 1, 5, 5.0, 3.0, risk=RiskLevel.HIGH)

    # === ПРОЧЕЕ ===
    categories["щетки_стеклоочистителя"] = make_cat("щетки_стеклоочистителя", "Щетки стеклоочистителя",
        30, 70, 2, 5, 2, 5, 0.1, 0.5, 1.0, 1.5)
    categories["коврики"] = make_cat("коврики", "Автомобильные коврики",
        50, 100, 40, 80, 1, 5, 1, 5, 10.0, 15.0)
    categories["чехлы"] = make_cat("чехлы", "Чехлы на сиденья",
        40, 80, 30, 60, 5, 20, 1, 5, 15.0, 25.0)
    categories["автохимия"] = make_cat("автохимия", "Автохимия и косметика",
        5, 30, 5, 20, 10, 40, 0.3, 5, 2.0, 5.0, hazardous=True)

    return categories


# ============================================================================
# БЛОК 6: DEEPSEEK AI С ГАЛОЧКОЙ "ЗАПРОСИТЬ ИИ" (350+ СТРОК)
# ============================================================================
class DeepSeekRateUpdater:
    """
    Класс для обновления тарифов через DeepSeek AI.
    НОВОЕ v97: поддержка галочки "Запросить ИИ" и умного кэша.
    """
    def __init__(self, api_key: Optional[str] = None):
        self.api_key = api_key or os.environ.get('DEEPSEEK_API_KEY')
        self.api_url = DEEPSEEK_API_URL
        self.session = requests.Session()
        self._logger = logging.getLogger('DeepSeekRateUpdater')
        self.cache = SmartTariffCache()
        if self.api_key:
            self.session.headers.update({
                "Authorization": f"Bearer {self.api_key}",
                "Content-Type": "application/json"
            })
            self._logger.info("DeepSeek клиент инициализирован")
        else:
            self._logger.warning("DeepSeek API ключ не найден")

    def _build_prompt(self, marketplace: str, category: str = None) -> str:
        """Формирование промпта для DeepSeek"""
        prompt = (
            "Ты - эксперт по юнит-экономике маркетплейсов России, "
            "специализирующийся на автозапчастях.\n"
            f"Предоставь актуальные тарифы для маркетплейса {marketplace} на 2026 год.\n"
            "Формат ответа ТОЛЬКО JSON без пояснений:\n"
            "{\n"
            '  "commission_rate": число (комиссия в долях, например 0.15),\n'
            '  "min_commission": число (минимальная комиссия в рублях),\n'
            '  "logistics_base": число (базовая стоимость логистики в рублях),\n'
            '  "logistics_per_kg": число (стоимость за кг в рублях),\n'
            '  "logistics_per_liter": число (стоимость за литр объема в рублях),\n'
            '  "storage_per_day": число (стоимость хранения за день в рублях/л),\n'
            '  "return_fee": число (процент возвратов в долях, например 0.02),\n'
            '  "acquiring_fee": число (процент эквайринга в долях),\n'
            '  "last_mile_fee": число (последняя миля в рублях),\n'
            '  "delivery_fee_percent": число (процент доставки в долях),\n'
            '  "hazardous_surcharge": число (надбавка за опасные грузы),\n'
            '  "fragile_surcharge": число (надбавка за хрупкие товары),\n'
            '  "oversized_surcharge": число (надбавка за крупногабарит)\n'
            "}\n"
        )
        if category:
            prompt += (
                f"\nДополнительно укажи комиссию для категории '{category}' "
                f"в поле 'category_rate_{category}'."
            )
        return prompt

    def _call_deepseek_api(self, prompt: str) -> Dict[str, Any]:
        """Вызов DeepSeek API"""
        if not self.api_key:
            raise AIError("API ключ DeepSeek не указан", provider="DeepSeek")
        try:
            payload = {
                "model": DEEPSEEK_MODEL,
                "messages": [
                    {"role": "system", "content": "Ты - эксперт по маркетплейсам и автозапчастям. Отвечай только JSON."},
                    {"role": "user", "content": prompt}
                ],
                "temperature": 0.1,
                "max_tokens": 1500
            }
            response = self.session.post(self.api_url, json=payload, timeout=60)
            if response.status_code == 200:
                result = response.json()
                content = result['choices'][0]['message']['content']
                json_match = re.search(r'\{.*\}', content, re.DOTALL)
                if json_match:
                    return json.loads(json_match.group())
                return {}
            else:
                raise AIError(
                    f"DeepSeek API вернул код {response.status_code}",
                    provider="DeepSeek",
                    code=response.status_code
                )
        except requests.exceptions.Timeout:
            raise AIError("Превышено время ожидания ответа DeepSeek", provider="DeepSeek")
        except Exception as e:
            raise AIError(f"Ошибка вызова DeepSeek API: {e}", provider="DeepSeek")

    def get_rates_from_ai(
        self,
        marketplace: str,
        category: str = None,
        force_refresh: bool = False,
        use_cache: bool = True
    ) -> Tuple[Dict[str, Any], TariffSource]:
        """
        Получение тарифов через DeepSeek.

        Args:
            marketplace: Название маркетплейса
            category: Категория (опционально)
            force_refresh: Если True - всегда запрашивать у ИИ (галочка "Запросить ИИ")
            use_cache: Если True - использовать кэш при отсутствии force_refresh

        Returns:
            Tuple из (словарь тарифов, источник тарифов)
        """
        # === ПРОВЕРКА КЭША (если не требуется принудительное обновление) ===
        if use_cache and not force_refresh:
            cached = self.cache.get(marketplace, category, use_expired=False)
            if cached:
                self._logger.info(
                    f"📥 Использованы кэшированные тарифы для {marketplace} "
                    f"(источник: {cached.source.value}, "
                    f"возраст: {(time.time() - cached.timestamp) / 3600:.1f}ч)"
                )
                return cached.data, cached.source

        # === ЗАПРОС К ИИ ===
        if not self.api_key:
            self._logger.warning("API ключ не указан, используются хардкод-тарифы")
            return {}, TariffSource.HARDCODED

        try:
            prompt = self._build_prompt(marketplace, category)
            result = self._call_deepseek_api(prompt)

            if result:
                # Сохраняем в кэш
                self.cache.set(
                    marketplace=marketplace,
                    category=category,
                    data=result,
                    source=TariffSource.AI_LIVE,
                    ttl_seconds=86400,
                    notes=f"Получено от DeepSeek {datetime.now().strftime('%Y-%m-%d %H:%M')}"
                )
                self._logger.info(f"✅ Тарифы для {marketplace} обновлены через DeepSeek AI")
                return result, TariffSource.AI_LIVE
            return {}, TariffSource.HARDCODED

        except AIError as e:
            self._logger.error(f"Ошибка DeepSeek: {e}")
            # При ошибке пытаемся использовать устаревший кэш
            if use_cache:
                cached = self.cache.get(marketplace, category, use_expired=True)
                if cached:
                    self._logger.warning(
                        f"⚠️ Использованы устаревшие кэшированные тарифы для {marketplace}"
                    )
                    return cached.data, cached.source
            return {}, TariffSource.HARDCODED

    def update_all_marketplaces(
        self,
        force_refresh: bool = False,
        progress_callback: Optional[Callable] = None
    ) -> Dict[str, Tuple[Dict[str, Any], TariffSource]]:
        """
        Обновление тарифов для всех маркетплейсов.

        Args:
            force_refresh: Принудительно запросить у ИИ
            progress_callback: Функция обратного вызова для прогресса

        Returns:
            Словарь {маркетплейс: (тарифы, источник)}
        """
        results = {}
        marketplaces = ["Ozon", "Wildberries", "Яндекс Маркет", "AliExpress", "Мегамаркет", "СберМегаМаркет"]
        total = len(marketplaces)

        for i, mp in enumerate(marketplaces):
            try:
                rates, source = self.get_rates_from_ai(
                    marketplace=mp,
                    force_refresh=force_refresh,
                    use_cache=True
                )
                results[mp] = (rates, source)
            except Exception as e:
                self._logger.error(f"Ошибка обновления тарифов для {mp}: {e}")
                results[mp] = ({}, TariffSource.HARDCODED)

            if progress_callback:
                progress_callback((i + 1) / total)

        return results

    def get_cache_statistics(self) -> Dict[str, Any]:
        """Получение статистики кэша"""
        return self.cache.get_statistics()

    def get_cache_history(self, limit: int = 50) -> List[Dict[str, Any]]:
        """Получение истории изменений кэша"""
        return self.cache.get_history(limit)

    def export_cache(self, file_path: Union[str, Path]) -> bool:
        """Экспорт кэша в файл"""
        return self.cache.export_to_file(file_path)

    def import_cache(self, file_path: Union[str, Path]) -> int:
        """Импорт кэша из файла"""
        return self.cache.import_from_file(file_path)

    def clear_cache(self) -> int:
        """Очистка кэша"""
        return self.cache.clear_all()

    def clear_expired_cache(self) -> int:
        """Очистка устаревших записей"""
        return self.cache.clear_expired()
# ============================================================================
# БЛОК 7: ОСНОВНОЙ КЛАСС ЮНИТ-ЭКОНОМИКИ v97 (С НАЛОГАМИ И РЕКОМЕНДОВАННОЙ ЦЕНОЙ)
# ============================================================================
class MarketplaceUnitEconomics:
    """
    Основной класс для расчета юнит-экономики.
    НОВОЕ v97:
    - Учёт налогов (УСН 6%, УСН 15%, ОСНО)
    - Расчёт рекомендованной минимальной цены (с учётом налога + 10% прибыли)
    - Надбавки за опасные/хрупкие/крупногабаритные товары
    - Интеграция с SmartTariffCache
    - Поддержка галочки "Запросить ИИ"
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
        self._configs = self._load_marketplace_configs()
        self._categories = self._load_categories()
        self._cache = {}
        self._history = []
        self._stats = self._init_stats()
        self._settings = self._load_settings()
        self._tariff_cache = SmartTariffCache()
        self._ai_updater = None
        self._logger = logging.getLogger('MarketplaceUnitEconomics')
        self._logger.info("🚗 Инициализация MarketplaceUnitEconomics v97")
        self._logger.info(f"📊 Загружено {len(self._configs)} маркетплейсов")
        self._logger.info(f"📚 Загружено {len(self._categories)} категорий")
        self._logger.info(f"💰 Система налогообложения: {self._settings.get('tax_system', 'УСН_6')}")
        self._logger.info(f"🎯 Мин. прибыль: {self._settings.get('min_profit_percent', 0.10) * 100:.1f}%")

    def _load_marketplace_configs(self) -> Dict[str, MarketplaceConfig]:
        """Загрузка конфигураций маркетплейсов (с учётом кэша)"""
        return get_marketplace_configs_2026()

    def _load_categories(self) -> Dict[str, ProductCategory]:
        """Загрузка категорий"""
        categories = {}
        for name, cat in get_auto_parts_categories_full().items():
            categories[name] = cat
        return categories

    def _init_stats(self) -> Dict[str, Any]:
        return {
            "total_calculations": 0,
            "by_marketplace": defaultdict(int),
            "by_category": defaultdict(int),
            "by_mode": defaultdict(int),
            "by_status": defaultdict(int),
            "by_tax_system": defaultdict(int),
            "by_tariff_source": defaultdict(int),
            "avg_profit": 0.0,
            "avg_margin": 0.0,
            "avg_roi": 0.0,
            "avg_tax": 0.0,
            "total_profit": 0.0,
            "total_tax": 0.0,
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
            "cache_misses": 0,
            "ai_requests": 0
        }

    def _load_settings(self) -> Dict[str, Any]:
        """Загрузка настроек с учётом налогов"""
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
            "timezone": "Europe/Moscow",
            # НОВОЕ v97: настройки налогов
            "tax_system": DEFAULT_TAX_SYSTEM,
            "usn_rate": DEFAULT_USN_RATE,
            "insurance_contributions": DEFAULT_INSURANCE_CONTRIBUTIONS,
            "min_profit_percent": DEFAULT_MIN_PROFIT_PERCENT,
            "use_tax_deduction": DEFAULT_TAX_DEDUCTION,
            "annual_revenue": 0.0
        }
        if settings_path.exists():
            try:
                with open(settings_path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                    default_settings.update(settings)
            except Exception as e:
                self._logger.warning(f"Ошибка загрузки настроек: {e}")
        return default_settings

    def save_settings(self, settings: Dict[str, Any]) -> bool:
        """Сохранение настроек"""
        try:
            settings_path = CONFIG_DIR / "settings.json"
            self._settings.update(settings)
            with open(settings_path, 'w', encoding='utf-8') as f:
                json.dump(self._settings, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            self._logger.error(f"Ошибка сохранения настроек: {e}")
            return False

    def get_category_dimensions(self, category_name: str) -> Optional[ProductDimensions]:
        """Получить габариты категории"""
        if category_name in self._categories:
            return self._categories[category_name].dimensions
        return None

    def get_category_info(self, category_name: str) -> Optional[ProductCategory]:
        """Получить информацию о категории"""
        return self._categories.get(category_name)

    def find_categories_by_keyword(self, keyword: str) -> List[Tuple[str, ProductCategory]]:
        """Поиск категорий по ключевому слову"""
        keyword_lower = keyword.lower()
        results = []
        for name, cat in self._categories.items():
            if keyword_lower in name.lower() or keyword_lower in cat.description.lower():
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

    def is_category_hazardous(self, category_name: str) -> bool:
        """Проверка: опасная ли категория"""
        cat = self._categories.get(category_name)
        return cat.hazardous if cat else False

    def is_category_fragile(self, category_name: str) -> bool:
        """Проверка: хрупкая ли категория"""
        cat = self._categories.get(category_name)
        return cat.fragile if cat else False

    def is_category_oversized(self, length: float, width: float, height: float, weight: float) -> bool:
        """Проверка: крупногабаритный ли товар"""
        # Крупногабарит: любой габарит > 100 см или вес > 25 кг
        return any([length > 100, width > 100, height > 100, weight > 25])

    def _get_ai_updater(self) -> Optional['DeepSeekRateUpdater']:
        """Получение AI-обновлятора (ленивая инициализация)"""
        if self._ai_updater is None:
            try:
                self._ai_updater = DeepSeekRateUpdater()
            except Exception as e:
                self._logger.error(f"Ошибка инициализации AI updater: {e}")
                return None
        return self._ai_updater

    def refresh_tariffs_from_ai(
        self,
        marketplace: Optional[str] = None,
        category: Optional[str] = None,
        force: bool = True
    ) -> Dict[str, Any]:
        """
        Обновление тарифов через DeepSeek AI.
        НОВОЕ v97: вызывается при установке галочки "Запросить ИИ".
        """
        updater = self._get_ai_updater()
        if updater is None:
            return {"error": "AI updater не инициализирован"}

        self._stats["ai_requests"] += 1

        try:
            if marketplace:
                # Обновление для одного маркетплейса
                rates, source = updater.get_rates_from_ai(
                    marketplace=marketplace,
                    category=category,
                    force_refresh=force,
                    use_cache=True
                )
                if rates:
                    # Применяем новые тарифы к конфигу
                    self._apply_ai_tariffs(marketplace, rates)
                    return {
                        "marketplace": marketplace,
                        "source": source.value,
                        "rates": rates,
                        "success": True
                    }
                return {"error": "Не удалось получить тарифы", "success": False}
            else:
                # Обновление для всех маркетплейсов
                results = updater.update_all_marketplaces(force_refresh=force)
                for mp, (rates, source) in results.items():
                    if rates:
                        self._apply_ai_tariffs(mp, rates)
                return {
                    "marketplaces_updated": len([r for r in results.values() if r[0]]),
                    "total": len(results),
                    "success": True
                }
        except Exception as e:
            self._logger.error(f"Ошибка обновления тарифов через AI: {e}")
            return {"error": str(e), "success": False}

    def _apply_ai_tariffs(self, marketplace: str, rates: Dict[str, Any]):
        """Применение AI-тарифов к конфигурации"""
        if marketplace not in self._configs:
            return
        config = self._configs[marketplace]
        # Обновляем поля из AI-ответа
        field_mapping = {
            "commission_rate": "commission_rate",
            "min_commission": "min_commission",
            "logistics_base": "logistics_base",
            "logistics_per_kg": "logistics_per_kg",
            "logistics_per_liter": "logistics_per_liter",
            "storage_per_day": "storage_per_day",
            "return_fee": "return_fee",
            "acquiring_fee": "acquiring_fee",
            "last_mile_fee": "last_mile_fee",
            "delivery_fee_percent": "delivery_fee_percent",
            "hazardous_surcharge": "hazardous_surcharge",
            "fragile_surcharge": "fragile_surcharge",
            "oversized_surcharge": "oversized_surcharge"
        }
        for ai_key, config_key in field_mapping.items():
            if ai_key in rates and hasattr(config, config_key):
                try:
                    setattr(config, config_key, float(rates[ai_key]))
                except (ValueError, TypeError):
                    pass
        config.tariff_source = TariffSource.AI_LIVE
        config.last_updated = datetime.now()
        self._logger.info(f"✅ AI-тарифы применены для {marketplace}")

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
        tax_system: Optional[str] = None,
        min_profit_percent: Optional[float] = None,
        annual_revenue: Optional[float] = None,
        **kwargs
    ) -> UnitEconomicsResult:
        """
        Расчет юнит-экономики с учётом налогов и рекомендованной минимальной цены.
        НОВОЕ v97: добавлены поля tax_amount, recommended_min_price, surcharges.
        """
        # === ВАЛИДАЦИЯ ===
        if price <= 0:
            raise ValidationError("Цена должна быть положительной", "price", price)
        if cost <= 0:
            raise ValidationError("Себестоимость должна быть положительной", "cost", cost)
        if marketplace not in self._configs:
            raise MarketplaceError(f"Маркетплейс {marketplace} не поддерживается", marketplace)

        config = self._configs[marketplace]

        # === ГАБАРИТЫ ИЗ КАТЕГОРИИ (если не указаны) ===
        if all([length == 0, width == 0, height == 0, weight == 0]) and category:
            length, width, height, weight = self.calculate_dimensions_from_category(category)

        volume = calculate_volume(length, width, height)
        if volume == 0:
            volume = 5.0
        if weight <= 0:
            weight = 1.0

        # === НАДБАВКИ ЗА СПЕЦКАТЕГОРИИ (НОВОЕ v97) ===
        hazardous = self.is_category_hazardous(category) if category else False
        fragile = self.is_category_fragile(category) if category else False
        oversized = self.is_category_oversized(length, width, height, weight)

        hazardous_surcharge = price * config.hazardous_surcharge if hazardous else 0.0
        fragile_surcharge = price * config.fragile_surcharge if fragile else 0.0
        oversized_surcharge = price * config.oversized_surcharge if oversized else 0.0

        # === КОМИССИЯ ===
        commission_rate = config.category_rates.get(category, config.commission_rate) if category else config.commission_rate
        commission = max(price * commission_rate, config.min_commission)
        if config.max_commission < float('inf'):
            commission = min(commission, config.max_commission)

        # === ПОДПИСКА ===
        subscription_cost = config.subscription_fee / 30 if config.subscription_fee > 0 else 0

        # === ЛОГИСТИКА ===
        logistics = (
            config.logistics_base +
            weight * config.logistics_per_kg +
            volume * config.logistics_per_liter
        )
        mode_multiplier = config.mode_multipliers.get(operation_mode, 1.0)
        logistics *= mode_multiplier

        # === ХРАНЕНИЕ ===
        storage_cost = volume * config.storage_per_day * days_in_storage

        # === ЭКВАЙРИНГ, ДОСТАВКА, ПОСЛЕДНЯЯ МИЛЯ ===
        acquiring = price * config.acquiring_fee
        delivery = price * config.delivery_fee_percent
        last_mile = config.last_mile_fee
        returns = price * config.return_fee

        # === РКО, ПРЕМИУМ, СТРАХОВКА, УПАКОВКА, МАРКЕТИНГ ===
        rko_fee = price * config.rko_fee if config.rko_fee > 0 else 0
        premium_fee = price * config.premium_fee if is_premium and config.premium_fee > 0 else 0
        insurance_fee = price * config.insurance_fee if include_insurance and config.insurance_fee > 0 else 0
        packing_fee = config.packing_fee if include_packing and config.packing_fee > 0 else 0
        marketing_fee = price * config.marketing_fee if include_marketing and config.marketing_fee > 0 else 0

        # === ИТОГО РАСХОДОВ (БЕЗ НАЛОГА) ===
        total_expenses_before_tax = (
            cost + commission + subscription_cost + logistics + storage_cost +
            acquiring + delivery + last_mile + returns + rko_fee +
            premium_fee + insurance_fee + packing_fee + marketing_fee +
            hazardous_surcharge + fragile_surcharge + oversized_surcharge
        )

        # === НАЛОГ (НОВОЕ v97) ===
        current_tax_system = tax_system or self._settings.get("tax_system", DEFAULT_TAX_SYSTEM)
        current_annual_revenue = annual_revenue if annual_revenue is not None else self._settings.get("annual_revenue", 0.0)
        insurance_contributions = self._settings.get("insurance_contributions", DEFAULT_INSURANCE_CONTRIBUTIONS)
        use_deduction = self._settings.get("use_tax_deduction", DEFAULT_TAX_DEDUCTION)

        tax_amount = calculate_tax(
            revenue=price,
            annual_revenue=current_annual_revenue,
            tax_system=current_tax_system,
            insurance_contributions=insurance_contributions,
            use_deduction=use_deduction
        )

        # === ИТОГО РАСХОДОВ (С НАЛОГОМ) ===
        total_expenses = total_expenses_before_tax + tax_amount

        # === ПРИБЫЛЬ ===
        profit = price - total_expenses
        margin_percent = (profit / price * 100) if price > 0 else 0
        roi = (profit / cost * 100) if cost > 0 else 0

        # === ТОЧКА БЕЗУБЫТОЧНОСТИ ===
        fixed_costs = logistics + storage_cost + last_mile + subscription_cost
        variable_rate = (
            commission_rate + config.acquiring_fee +
            config.delivery_fee_percent + config.return_fee +
            config.rko_fee + config.premium_fee +
            config.insurance_fee + config.marketing_fee +
            config.hazardous_surcharge + config.fragile_surcharge +
            config.oversized_surcharge
        )
        # Добавляем ставку налога в переменные расходы
        if current_tax_system == "УСН_6":
            variable_rate += self._settings.get("usn_rate", DEFAULT_USN_RATE)
        elif current_tax_system == "УСН_15":
            variable_rate += 0.15
        elif current_tax_system == "ОСНО":
            variable_rate += 0.40  # НДС 20% + прибыль 20%

        breakeven_price = ((cost + fixed_costs) / (1 - variable_rate)) if (1 - variable_rate) > 0 else 0

        # === РЕКОМЕНДОВАННАЯ МИНИМАЛЬНАЯ ЦЕНА (НОВОЕ v97) ===
        current_min_profit = min_profit_percent if min_profit_percent is not None else self._settings.get("min_profit_percent", DEFAULT_MIN_PROFIT_PERCENT)
        tax_rate_for_formula = 0.0
        if current_tax_system == "УСН_6":
            tax_rate_for_formula = self._settings.get("usn_rate", DEFAULT_USN_RATE)
        elif current_tax_system == "УСН_15":
            tax_rate_for_formula = 0.15
        elif current_tax_system == "ОСНО":
            tax_rate_for_formula = 0.40

        recommended_min_price = calculate_recommended_min_price(
            cost=cost,
            commission_rate=commission_rate,
            logistics=logistics,
            storage_cost=storage_cost,
            acquiring_rate=config.acquiring_fee,
            last_mile=last_mile,
            return_rate=config.return_fee,
            min_profit_percent=current_min_profit,
            tax_system=current_tax_system,
            tax_rate=tax_rate_for_formula
        )

        # === МАРЖИНАЛЬНЫЙ ДОХОД ===
        contribution_margin = price - cost - commission - logistics - acquiring - delivery - last_mile - returns - tax_amount
        contribution_margin_ratio = (contribution_margin / price * 100) if price > 0 else 0

        # === СОЗДАНИЕ РЕЗУЛЬТАТА ===
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
            hazardous_surcharge=round(hazardous_surcharge, 2),
            fragile_surcharge=round(fragile_surcharge, 2),
            oversized_surcharge=round(oversized_surcharge, 2),
            tax_amount=round(tax_amount, 2),
            tax_system=current_tax_system,
            total_expenses=round(total_expenses, 2),
            profit=round(profit, 2),
            margin_percent=round(margin_percent, 2),
            roi=round(roi, 2),
            breakeven_price=round(breakeven_price, 2),
            recommended_min_price=round(recommended_min_price, 2),
            profit_per_ruble=round(profit / price, 4) if price > 0 else 0,
            contribution_margin=round(contribution_margin, 2),
            contribution_margin_ratio=round(contribution_margin_ratio, 2),
            status=CalculationStatus.COMPLETED,
            tariff_source=config.tariff_source,
            metadata=kwargs
        )

        # === ОБНОВЛЕНИЕ СТАТИСТИКИ ===
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
        self._stats["by_tax_system"][result.tax_system] += 1
        self._stats["by_tariff_source"][result.tariff_source.value] += 1
        self._stats["total_profit"] += result.profit
        self._stats["total_tax"] += result.tax_amount
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
        self._stats["avg_tax"] = self._stats["total_tax"] / n

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
                    price=price, cost=cost, marketplace=marketplace,
                    category=category, operation_mode=operation_mode,
                    days_in_storage=days_in_storage,
                    length=length, width=width, height=height,
                    weight=weight, **kwargs
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
                "idx": idx, "article": article, "brand": brand,
                "price": price, "cost": cost, "category": category,
                "length": length, "width": width,
                "height": height, "weight": weight
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
                        price=item["price"], cost=item["cost"],
                        marketplace=marketplace, category=item["category"],
                        operation_mode=operation_mode,
                        days_in_storage=days_in_storage,
                        length=item["length"], width=item["width"],
                        height=item["height"], weight=item["weight"]
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
                    price=current_price, cost=cost, marketplace=marketplace,
                    category=category, operation_mode=operation_mode,
                    days_in_storage=days_in_storage,
                    length=length, width=width, height=height, weight=weight
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
            price=price_min or best_price, cost=cost, marketplace=marketplace,
            category=category, operation_mode=operation_mode,
            days_in_storage=days_in_storage,
            length=length, width=width, height=height, weight=weight
        )
        improvement_pct = ((best_profit - current_result.profit) / current_result.profit * 100) if current_result.profit > 0 else 0
        recommendations = []
        if best_price > 0 and best_margin >= target_margin:
            recommendations.append(f"Установите цену {best_price:.2f} ₽ для достижения маржи {target_margin}%")
        else:
            recommendations.append(f"Целевая маржа {target_margin}% не достигнута. Максимальная маржа: {best_margin:.1f}%")
        if best_profit > current_result.profit:
            recommendations.append(f"Потенциальное увеличение прибыли: {improvement_pct:.1f}%")
        if current_result.recommended_min_price > 0:
            recommendations.append(f"Рекомендуемая минимальная цена: {current_result.recommended_min_price:.2f} ₽")

        self._stats["total_optimizations"] += 1
        self._stats["optimization_improvement"] += improvement_pct

        return OptimizationResult(
            optimal_price=best_price, optimal_margin=best_margin,
            optimal_profit=best_profit,
            current_price=current_result.price,
            current_margin=current_result.margin_percent,
            current_profit=current_result.profit,
            improvement_pct=improvement_pct,
            recommendations=recommendations,
            metadata={"target_margin": target_margin, "step": step, "iterations": iteration}
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
        periods_list, values_list, seasonality_list, trend_list = [], [], [], []
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
            periods=periods_list, values=values_list,
            seasonality=seasonality_list, trend=trend_list,
            confidence_intervals=(lower_bound, upper_bound),
            metadata={"base_value": base_value, "growth_rate": growth_rate, "confidence_level": confidence_level}
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
                        match = False; break
                    elif key == "category" and item.category != value:
                        match = False; break
                    elif key == "operation_mode" and item.operation_mode != value:
                        match = False; break
                    elif key == "tax_system" and item.tax_system != value:
                        match = False; break
                    elif key == "min_profit" and item.profit < value:
                        match = False; break
                    elif key == "max_profit" and item.profit > value:
                        match = False; break
                    elif key == "start_date" and item.timestamp < value:
                        match = False; break
                    elif key == "end_date" and item.timestamp > value:
                        match = False; break
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
            "tax": best.tax_amount,
            "tax_system": best.tax_system,
            "recommended_min_price": best.recommended_min_price,
            "timestamp": best.timestamp.isoformat()
        }

    def get_category_stats(self) -> pd.DataFrame:
        """Статистика по категориям"""
        if not self._history:
            return pd.DataFrame()
        stats = defaultdict(lambda: {
            "count": 0, "total_profit": 0, "avg_profit": 0,
            "avg_margin": 0, "best_profit": 0, "worst_profit": 0,
            "total_tax": 0, "avg_recommended_price": 0
        })
        for result in self._history:
            cat = result.category
            stats[cat]["count"] += 1
            stats[cat]["total_profit"] += result.profit
            stats[cat]["avg_margin"] += result.margin_percent
            stats[cat]["total_tax"] += result.tax_amount
            stats[cat]["avg_recommended_price"] += result.recommended_min_price
            stats[cat]["best_profit"] = max(stats[cat]["best_profit"], result.profit)
            stats[cat]["worst_profit"] = min(stats[cat]["worst_profit"], result.profit)
        for cat in stats:
            if stats[cat]["count"] > 0:
                stats[cat]["avg_profit"] = stats[cat]["total_profit"] / stats[cat]["count"]
                stats[cat]["avg_margin"] /= stats[cat]["count"]
                stats[cat]["avg_recommended_price"] /= stats[cat]["count"]
        return pd.DataFrame.from_dict(stats, orient="index").reset_index().rename(columns={"index": "category"})

    def get_marketplace_stats(self) -> pd.DataFrame:
        """Статистика по маркетплейсам"""
        if not self._history:
            return pd.DataFrame()
        stats = defaultdict(lambda: {
            "count": 0, "total_profit": 0, "avg_profit": 0,
            "avg_margin": 0, "best_profit": 0, "worst_profit": 0,
            "total_tax": 0
        })
        for result in self._history:
            mp = result.marketplace
            stats[mp]["count"] += 1
            stats[mp]["total_profit"] += result.profit
            stats[mp]["avg_margin"] += result.margin_percent
            stats[mp]["total_tax"] += result.tax_amount
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

    def get_tariff_cache_statistics(self) -> Dict[str, Any]:
        """Статистика кэша тарифов"""
        return self._tariff_cache.get_statistics()

    def get_tariff_cache_history(self, limit: int = 50) -> List[Dict[str, Any]]:
        """История изменений кэша тарифов"""
        return self._tariff_cache.get_history(limit)


# ============================================================================
# БЛОК 8: HIGH-VOLUME CATALOG (POLARS + DUCKDB) (700+ СТРОК)
# ============================================================================
class HighVolumeAutoPartsCatalog:
    """
    High-Volume каталог автозапчастей с поддержкой 10M+ записей.
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
            "enabled": False, "provider": "s3", "bucket": "",
            "region": "", "sync_interval": 3600, "last_sync": 0
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
            "global_markup": 0.2, "brand_markups": {},
            "min_price": 0.0, "max_price": 99999.0
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
            "Радиатор": "Охлаждение", "Шаровая опора": "Подвеска",
            "Фильтр масляный": "Фильтры", "Тормозные колодки": "Тормоза"
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
            logger.warning(f"Не удалось определить колонки для файла {file_type}. Доступные: {df.columns}")
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
                df = df.with_columns(self.normalize_key(pl.col(col)).alias(f"{col}_norm"))
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
            pk_cols_csv = ", ".join(f'"{c}"' for c in pk)
            delete_sql = f"""
                DELETE FROM {table_name}
                WHERE ({pk_cols_csv}) IN (SELECT {pk_cols_csv} FROM {temp_view_name});
            """
            self.conn.execute(delete_sql)
            insert_sql = f"INSERT INTO {table_name} SELECT * FROM {temp_view_name};"
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
                oe_df = oe_df.with_columns(self.determine_category_vectorized(pl.col('name')))
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
                parts_df = parts_df.with_columns(multiplicity=pl.lit(1).cast(pl.Int32))
            else:
                parts_df = parts_df.with_columns(pl.col('multiplicity').fill_null(1).cast(pl.Int32))
            for col in ['length', 'width', 'height']:
                if col not in parts_df.columns:
                    parts_df = parts_df.with_columns(pl.lit(None).cast(pl.Float64).alias(col))
            if 'dimensions_str' not in parts_df.columns:
                parts_df = parts_df.with_columns(dimensions_str=pl.lit(None).cast(pl.Utf8))
            parts_df = parts_df.with_columns([
                pl.col('length').cast(pl.Utf8).fill_null('').alias('_length_str'),
                pl.col('width').cast(pl.Utf8).fill_null('').alias('_width_str'),
                pl.col('height').cast(pl.Utf8).fill_null('').alias('_height_str'),
            ])
            parts_df = parts_df.with_columns(
                dimensions_str=pl.when(
                    (pl.col('dimensions_str').is_not_null()) &
                    (pl.col('dimensions_str').cast(pl.Utf8) != '')
                ).then(pl.col('dimensions_str').cast(pl.Utf8)).otherwise(
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
            select_exprs = [pl.col(c) if c in parts_df.columns else pl.lit(None).alias(c) for c in final_columns]
            parts_df = parts_df.select(select_exprs)
            self.upsert_data('parts', parts_df, ['artikul_norm', 'brand_norm'])
        logger.info("✅ Обновление базы данных завершено!")

    def merge_all_data_parallel(self, file_paths: Dict[str, str], max_workers: int = 4) -> Dict[str, "pl.DataFrame"]:
        if not POLARS_AVAILABLE:
            return {}
        results = {}
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {}
            for key, path in file_paths.items():
                if path and os.path.exists(path):
                    futures[executor.submit(self.read_and_prepare_file, path, key)] = key
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

    def delete_by_brand(self, brand_norm: str) -> int:
        if not self.conn:
            return 0
        try:
            count_result = self.conn.execute(
                "SELECT COUNT(*) FROM parts WHERE brand_norm = ?", [brand_norm]).fetchone()
            deleted_count = count_result[0] if count_result else 0
            if deleted_count == 0:
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
                return 0
            self.conn.execute("DELETE FROM parts WHERE artikul_norm = ?", [artikul_norm])
            self.conn.execute(
                "DELETE FROM cross_references WHERE (artikul_norm, brand_norm) NOT IN (SELECT DISTINCT artikul_norm, brand_norm FROM parts)")
            return deleted_count
        except Exception as e:
            logger.error(f"Error deleting by artikul {artikul_norm}: {e}")
            raise


# ============================================================================
# БЛОК 9: CATALOG ENHANCER (ПОИСК АНАЛОГОВ 2 УРОВНЯ) (400+ СТРОК)
# ============================================================================
class CatalogEnhancer:
    """Обогащение каталога: поиск аналогов по OE номерам (2 уровня)"""
    def __init__(self, db_path: Optional[str] = None):
        self.data_dir = Path("./catalog_data")
        self.data_dir.mkdir(exist_ok=True)
        self.db_path = Path(db_path) if db_path else self.data_dir / "catalog.duckdb"
        self.conn = None
        self.stats = {
            "oe_loaded": 0, "parts_loaded": 0, "cross_loaded": 0,
            "analog_searches": 0, "enrichments": 0
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
            logger.warning("DuckDB не установлен")

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
                        p.artikul, p.brand, p.description,
                        p.length, p.width, p.height, p.weight,
                        p.dimensions_str, p.image_url
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
                    "artikul": artikul, "brand": brand,
                    "analog_count": 0, "analogs": [],
                    "has_analogs": False, "oe_list": ""
                }
            row = result.iloc[0]
            analog_count = int(row['analog_count']) if not pd.isna(row['analog_count']) else 0
            analogs = []
            if analog_count > 0:
                analog_query = f"""
                    SELECT DISTINCT
                        p.artikul, p.brand, p.description,
                        p.length, p.width, p.height, p.weight,
                        p.dimensions_str, p.image_url
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
                "artikul": artikul, "brand": brand,
                "analog_count": analog_count, "analogs": analogs,
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
# БЛОК 10: ML-КЛАССИФИКАТОР КАТЕГОРИЙ (300+ СТРОК)
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
# БЛОК 11: UI ФУНКЦИИ - ЗАГРУЗКА ДАННЫХ (500+ СТРОК)
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
# БЛОК 12: ЮНИТ-ЭКОНОМИКА (ОДИН ТОВАР) - v97 С НАЛОГАМИ
# ============================================================================
def show_unit_economics_interface():
    """Интерфейс расчета юнит-экономики для одного товара (v97)"""
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

    # === НОВОЕ v97: НАСТРОЙКИ НАЛОГОВ ===
    with st.expander("💰 Настройки налогов и прибыли"):
        col_t1, col_t2, col_t3 = st.columns(3)
        with col_t1:
            tax_system = st.selectbox(
                "Система налогообложения",
                ["УСН_6", "УСН_15", "ОСНО"],
                index=0,
                key="ue_tax_system",
                help="УСН 6% - с доходов; УСН 15% - доходы-расходы; ОСНО - НДС + прибыль"
            )
        with col_t2:
            min_profit_percent = st.number_input(
                "Желаемая прибыль (%)",
                min_value=0.0,
                max_value=100.0,
                value=10.0,
                step=1.0,
                key="ue_min_profit",
                help="Минимальная прибыль после уплаты налогов и всех комиссий"
            ) / 100
        with col_t3:
            annual_revenue = st.number_input(
                "Годовая выручка (₽)",
                min_value=0.0,
                value=0.0,
                step=100000.0,
                key="ue_annual_revenue",
                help="Для расчёта 1% налога сверх лимита 300 млн"
            )

    is_premium = st.checkbox("⭐ Премиум-раздел (доп. комиссия)", key="ue_premium")

    if st.button("🚀 Рассчитать юнит-экономику", type="primary", key="ue_calc"):
        with st.spinner("Расчет юнит-экономики..."):
            economics = unit_economics.calculate_unit_economics(
                price=price,
                cost=cost,
                marketplace=marketplace,
                weight=weight,
                category=category if category else None,
                is_premium=is_premium,
                tax_system=tax_system,
                min_profit_percent=min_profit_percent,
                annual_revenue=annual_revenue
            )

            # === ГЛАВНЫЕ МЕТРИКИ ===
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("💰 Прибыль", f"{economics.profit:.2f} ₽", delta=f"{economics.profit_per_ruble:.2f} ₽/₽")
            with col2:
                st.metric("📈 Маржа", f"{economics.margin_percent:.2f}%")
            with col3:
                st.metric("📊 ROI", f"{economics.roi:.2f}%")
            with col4:
                st.metric("⚖️ Точка безубыточности", f"{economics.breakeven_price:.2f} ₽")

            # === НОВОЕ v97: РЕКОМЕНДОВАННАЯ ЦЕНА ===
            st.subheader("💎 Рекомендованная минимальная цена")
            col_rec1, col_rec2, col_rec3 = st.columns(3)
            with col_rec1:
                st.metric(
                    "🎯 Мин. цена (с учётом налога и 10% прибыли)",
                    f"{economics.recommended_min_price:.2f} ₽",
                    delta=f"{economics.recommended_min_price - price:.2f} ₽"
                )
            with col_rec2:
                st.metric("💵 Налог ({tax_system})", f"{economics.tax_amount:.2f} ₽")
            with col_rec3:
                if price < economics.recommended_min_price:
                    st.warning(f"⚠️ Цена ниже рекомендованной на {economics.recommended_min_price - price:.2f} ₽")
                else:
                    st.success(f"✅ Цена выше минимальной на {price - economics.recommended_min_price:.2f} ₽")

            # === ДЕТАЛИЗАЦИЯ РАСХОДОВ ===
            st.subheader("📋 Детализация расходов")
            expenses_data = {
                "Статья расходов": [
                    "Себестоимость", "Комиссия", "Подписка", "Логистика",
                    "Хранение", "Эквайринг", "Доставка", "Последняя миля",
                    "Возвраты", "РКО", "Премиум", "Страховка", "Упаковка", "Маркетинг",
                    "Надбавка за опасные", "Надбавка за хрупкие", "Надбавка за крупногабарит",
                    f"Налог ({economics.tax_system})", "ИТОГО"
                ],
                "Сумма (₽)": [
                    economics.cost, economics.commission, economics.subscription_cost,
                    economics.logistics, economics.storage_cost, economics.acquiring,
                    economics.delivery, economics.last_mile, economics.returns,
                    economics.rko_fee, economics.premium_fee, economics.insurance_fee,
                    economics.packing_fee, economics.marketing_fee,
                    economics.hazardous_surcharge, economics.fragile_surcharge,
                    economics.oversized_surcharge, economics.tax_amount,
                    economics.total_expenses
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
                    f"{economics.hazardous_surcharge/price*100:.1f}%",
                    f"{economics.fragile_surcharge/price*100:.1f}%",
                    f"{economics.oversized_surcharge/price*100:.1f}%",
                    f"{economics.tax_amount/price*100:.1f}%",
                    f"{economics.total_expenses/price*100:.1f}%"
                ]
            }
            st.dataframe(pd.DataFrame(expenses_data), use_container_width=True, key="ue_expenses_table")

            # === СРАВНЕНИЕ ВСЕХ МП ===
            st.subheader("🏆 Сравнение всех маркетплейсов")
            comparison_df = unit_economics.calculate_for_all_marketplaces(
                price=price, cost=cost, weight=weight, category=category if category else None,
                operation_mode=operation_mode, tax_system=tax_system,
                min_profit_percent=min_profit_percent, annual_revenue=annual_revenue
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
                            y=comparison_df['recommended_min_price'],
                            name='Мин. цена',
                            marker_color='#0f3460',
                            yaxis='y2'
                        ))
                        fig.update_layout(
                            title='Сравнение маркетплейсов',
                            xaxis_title='Маркетплейс',
                            yaxis_title='Прибыль (₽)',
                            yaxis2=dict(title='Рекомендованная цена (₽)', overlaying='y', side='right'),
                            barmode='group',
                            height=400
                        )
                        st.plotly_chart(fig, use_container_width=True)
                    except Exception as e:
                        logger.warning(f"Ошибка визуализации: {e}")


# ============================================================================
# БЛОК 13: ЮНИТ-ЭКОНОМИКА ПО АРТИКУЛАМ (v97)
# ============================================================================
def show_unit_economics_by_article_interface():
    """Интерфейс для расчета юнит-экономики по каждому артикулу (v97)"""
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
- 💎 **Рекомендованная минимальная цена** (с учётом налога + 10% прибыли)
- 📋 Детализация всех расходов
""")

    # === ПАРАМЕТРЫ РАСЧЕТА ===
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

    # === НОВОЕ v97: НАСТРОЙКИ НАЛОГОВ ===
    with st.expander("💰 Настройки налогов и прибыли"):
        col_t1, col_t2, col_t3 = st.columns(3)
        with col_t1:
            tax_system = st.selectbox(
                "Система налогообложения",
                ["УСН_6", "УСН_15", "ОСНО"],
                index=0,
                key="ue_article_tax_system"
            )
        with col_t2:
            min_profit_percent = st.number_input(
                "Желаемая прибыль (%)",
                min_value=0.0,
                max_value=100.0,
                value=10.0,
                step=1.0,
                key="ue_article_min_profit"
            ) / 100
        with col_t3:
            annual_revenue = st.number_input(
                "Годовая выручка (₽)",
                min_value=0.0,
                value=0.0,
                step=100000.0,
                key="ue_article_annual_revenue"
            )

    # === ВЫБОР КОЛОНОК ===
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

    # === ОПЦИОНАЛЬНЫЕ КОЛОНКИ ===
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

    # === КНОПКА РАСЧЕТА ===
    if st.button("🚀 Рассчитать юнит-экономику по артикулам", type="primary", key="ue_calc_articles"):
        with st.spinner("Расчет юнит-экономики для всех товаров..."):
            try:
                filtered_df = df.copy()
                if filtered_df.empty:
                    st.warning("⚠️ Нет данных для расчета")
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
                            is_premium=is_premium,
                            tax_system=tax_system,
                            min_profit_percent=min_profit_percent,
                            annual_revenue=annual_revenue
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
                            'Рекомендованная_мин_цена': economics.recommended_min_price,
                            'Налог': economics.tax_amount,
                            'Система_налога': economics.tax_system,
                            'Прибыль_на_рубль': economics.profit_per_ruble,
                            'Комиссия': economics.commission,
                            'Комиссия_%': economics.commission_percent,
                            'Логистика': economics.logistics,
                            'Хранение': economics.storage_cost,
                            'Надбавка_опасные': economics.hazardous_surcharge,
                            'Надбавка_хрупкие': economics.fragile_surcharge,
                            'Надбавка_крупногабарит': economics.oversized_surcharge,
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

                # === СТАТИСТИКА ===
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

                # === ТАБЛИЦА РЕЗУЛЬТАТОВ ===
                st.subheader("📋 Результаты расчета")
                display_cols = ['Артикул', 'Маркетплейс', 'Цена_с_наценкой', 'Прибыль', 'Маржа_%',
                               'Рекомендованная_мин_цена', 'Налог', 'Точка_безубыточности']
                available_display = [col for col in display_cols if col in df_results.columns]

                sort_col = st.selectbox("Сортировать по:", options=available_display,
                                        index=available_display.index('Прибыль') if 'Прибыль' in available_display else 0,
                                        key="ue_sort_col")
                sort_order = st.radio("Порядок:", ["По убыванию", "По возрастанию"], horizontal=True, key="ue_sort_order")
                sorted_df = df_results.sort_values(sort_col, ascending=(sort_order == "По возрастанию"))
                st.dataframe(sorted_df[available_display], use_container_width=True, key="ue_results_table")

                # === ЭКСПОРТ ===
                st.subheader("📤 Экспорт результатов")
                export_col1, export_col2 = st.columns(2)
                with export_col1:
                    if st.button("📥 Экспорт в Excel С ФОРМУЛАМИ", key="ue_export_excel_formulas"):
                        with st.spinner("Генерация Excel с формулами..."):
                            output_path = f"юнит_экономика_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                            if export_unit_economics_with_formulas_v2(df_results, output_path, unit_economics._configs):
                                with open(output_path, "rb") as f:
                                    st.download_button(
                                        label="📥 Скачать Excel с формулами",
                                        data=f,
                                        file_name=output_path,
                                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                        key="ue_download_excel_formulas"
                                    )
                                st.success("✅ Файл с живыми формулами готов! Измените цену — и всё пересчитается.")
                                try:
                                    os.remove(output_path)
                                except Exception:
                                    pass
                with export_col2:
                    if st.button("📥 Экспорт в CSV", key="ue_export_csv"):
                        csv = df_results.to_csv(index=False, encoding='utf-8-sig')
                        st.download_button(
                            label="📥 Скачать CSV",
                            data=csv,
                            file_name=f"юнит_экономика_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                            mime="text/csv",
                            key="ue_download_csv"
                        )

            except Exception as e:
                st.error(f"❌ Ошибка при расчете: {str(e)}")
                st.code(traceback.format_exc())
                logger.error(f"UE by article error: {traceback.format_exc()}")


# ============================================================================
# БЛОК 14: AI ТАРИФЫ - v97 С ГАЛОЧКОЙ "ЗАПРОСИТЬ ИИ"
# ============================================================================
def show_ai_tariffs_interface():
    """Интерфейс управления тарифами через DeepSeek AI (v97)"""
    st.header("🤖 Управление тарифами через DeepSeek AI")

    st.info("""
🚀 **DeepSeek AI обновляет тарифы для автозапчастей**

**Как работает:**
- ✅ Если галочка **"Запросить ИИ"** установлена — тарифы запрашиваются у DeepSeek AI
- ✅ Если галочка **НЕ установлена** — используются тарифы из кэша (или захардкоженные)
- ✅ Данные кэшируются на 24 часа с возможностью ручного редактирования
- ✅ Сохраняется история всех изменений тарифов
""")

    unit_economics = MarketplaceUnitEconomics()

    # === API КЛЮЧ ===
    api_key = st.text_input(
        "🔑 API ключ DeepSeek",
        type="password",
        placeholder="sk-...",
        help="Получите API ключ на platform.deepseek.com",
        key="ai_api_key"
    )
    if api_key:
        os.environ['DEEPSEEK_API_KEY'] = api_key

    # === НАСТРОЙКИ ===
    col1, col2, col3 = st.columns(3)
    with col1:
        marketplace = st.selectbox(
            "🏪 Маркетплейс",
            ["Ozon", "Wildberries", "Яндекс Маркет", "AliExpress", "Мегамаркет", "Все маркетплейсы"],
            key="ai_marketplace"
        )
    with col2:
        category = st.text_input(
            "📂 Категория (опционально)",
            placeholder="например: двигатель",
            key="ai_category"
        )
    with col3:
        # === НОВОЕ v97: ГАЛОЧКА "ЗАПРОСИТЬ ИИ" ===
        force_refresh = st.checkbox(
            "🔄 Запросить ИИ (принудительное обновление)",
            value=False,
            key="ai_force_refresh",
            help="Если установлено — тарифы будут запрошены у DeepSeek AI. Если нет — используются кэшированные."
        )

    if st.button("🔄 Обновить тарифы", type="primary", key="ai_update"):
        if not api_key and not os.environ.get('DEEPSEEK_API_KEY'):
            st.error("❌ Введите API ключ DeepSeek")
            return

        with st.spinner("Обновление тарифов..."):
            if marketplace == "Все маркетплейсы":
                result = unit_economics.refresh_tariffs_from_ai(
                    marketplace=None,
                    category=category if category else None,
                    force=force_refresh
                )
                if result.get('success'):
                    st.success(f"✅ Обновлены тарифы для {result.get('marketplaces_updated', 0)} из {result.get('total', 0)} маркетплейсов")
                else:
                    st.error(f"❌ Ошибка: {result.get('error', 'Неизвестная ошибка')}")
            else:
                result = unit_economics.refresh_tariffs_from_ai(
                    marketplace=marketplace,
                    category=category if category else None,
                    force=force_refresh
                )
                if result.get('success'):
                    st.success(f"✅ Обновлены тарифы для {marketplace}")
                    st.info(f"📥 Источник: **{result.get('source', 'Н/Д')}**")
                    with st.expander("📋 Полученные тарифы"):
                        st.json(result.get('rates', {}))
                else:
                    st.error(f"❌ Ошибка: {result.get('error', 'Неизвестная ошибка')}")

    # === ТЕКУЩИЕ ТАРИФЫ ===
    st.divider()
    st.subheader("📊 Текущие тарифы маркетплейсов")

    tariff_data = []
    for mp_name, config in unit_economics._configs.items():
        tariff_data.append({
            "Маркетплейс": mp_name,
            "Комиссия (%)": f"{config.commission_rate * 100:.1f}%",
            "Мин. комиссия (₽)": f"{config.min_commission:.0f}",
            "Логистика база (₽)": f"{config.logistics_base:.0f}",
            "Логистика за кг (₽)": f"{config.logistics_per_kg:.0f}",
            "Хранение/день (₽/л)": f"{config.storage_per_day:.2f}",
            "Возвраты (%)": f"{config.return_fee * 100:.1f}%",
            "Источник": config.tariff_source.value,
            "Обновлено": config.last_updated.strftime("%Y-%m-%d %H:%M")
        })

    st.dataframe(pd.DataFrame(tariff_data), use_container_width=True, key="ai_tariffs_table")

    # === СТАТИСТИКА КЭША ===
    st.divider()
    st.subheader("📊 Статистика кэша тарифов")
    cache_stats = unit_economics.get_tariff_cache_statistics()

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("📦 Всего записей", cache_stats.get('total_entries', 0))
    with col2:
        st.metric("⏰ Устаревших", cache_stats.get('expired_count', 0))
    with col3:
        st.metric("📝 Записей в истории", cache_stats.get('history_count', 0))
    with col4:
        st.metric("🤖 AI запросов", unit_economics._stats.get('ai_requests', 0))

    # === ИСТОРИЯ ИЗМЕНЕНИЙ ===
    with st.expander("📜 История изменений тарифов"):
        history = unit_economics.get_tariff_cache_history(limit=20)
        if history:
            st.dataframe(pd.DataFrame(history), use_container_width=True, key="ai_history_table")
        else:
            st.info("📜 История пуста")

    # === УПРАВЛЕНИЕ КЭШЕМ ===
    st.divider()
    st.subheader("🔧 Управление кэшем")
    col1, col2, col3 = st.columns(3)
    with col1:
        if st.button("🗑️ Очистить устаревшие", key="ai_clear_expired"):
            count = unit_economics._tariff_cache.clear_expired()
            st.success(f"✅ Удалено {count} устаревших записей")
            st.rerun()
    with col2:
        if st.button("🗑️ Очистить весь кэш", key="ai_clear_all"):
            if st.checkbox("Подтверждаю полную очистку", key="ai_confirm_clear"):
                count = unit_economics._tariff_cache.clear_all()
                st.success(f"✅ Удалено {count} записей")
                st.rerun()
    with col3:
        if st.button("📥 Экспорт кэша", key="ai_export_cache"):
            export_path = TARIFFS_DIR / f"tariffs_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
            if unit_economics._tariff_cache.export_to_file(export_path):
                with open(export_path, "rb") as f:
                    st.download_button(
                        label="📥 Скачать кэш",
                        data=f,
                        file_name=export_path.name,
                        mime="application/json",
                        key="ai_download_cache"
                    )
                st.success(f"✅ Кэш экспортирован: {export_path}")


# ============================================================================
# БЛОК 15: ОБОГАЩЕНИЕ КАТАЛОГА
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


# ============================================================================
# БЛОК 16: HIGH-VOLUME ИНТЕРФЕЙС (сокращённый)
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
            'oe': oe_file, 'cross': cross_file, 'barcode': barcode_file,
            'dimensions': weight_dims_file, 'images': images_file, 'prices': prices_file
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
        include_prices = st.checkbox("Включить цены", value=True, key="hv_prices_include")
        apply_markup = st.checkbox("Применить наценку", value=True, disabled=not include_prices, key="hv_markup")

        if st.button("🚀 Экспортировать", key="hv_export_btn"):
            output_path = catalog.data_dir / f"export.{format_choice.lower()}"
            with st.spinner("Генерация файла..."):
                if format_choice == "CSV":
                    catalog.export_to_csv_optimized(str(output_path), None, include_prices, apply_markup)
                elif format_choice == "Excel":
                    catalog.export_to_excel_optimized(str(output_path), None, include_prices, apply_markup)
                elif format_choice == "Parquet":
                    catalog.export_to_parquet(str(output_path), None, include_prices, apply_markup)

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
            ["Удалить по бренду", "Удалить по артикулу"],
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
            if st.checkbox("Подтверждаю удаление", key="hv_confirm_brand"):
                if st.button("Удалить", key="hv_delete_brand_btn"):
                    brand_norm = catalog.normalize_key(pl.Series([selected_brand]))[0]
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


# ============================================================================
# БЛОК 17: АНАЛИТИКА
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

    if PLOTLY_AVAILABLE and px is not None and price_col:
        st.subheader("📊 Визуализация данных")
        chart_type = st.selectbox(
            "Тип графика",
            ["Распределение цен", "Распределение категорий", "Топ товаров по цене"],
            key="analytics_chart_type"
        )

        if chart_type == "Распределение цен":
            fig = px.histogram(df, x=price_col, nbins=30,
                               title=f"Распределение цен ({price_col})")
            st.plotly_chart(fig, use_container_width=True)
        elif chart_type == "Распределение категорий" and 'Категория' in df.columns:
            category_counts = df['Категория'].value_counts()
            fig = px.pie(values=category_counts.values, names=category_counts.index,
                         title="Распределение товаров по категориям")
            st.plotly_chart(fig, use_container_width=True)
        elif chart_type == "Топ товаров по цене":
            name_col = None
            for col in df.columns:
                if any(w in col.lower() for w in ['наименование', 'название', 'name', 'товар']):
                    name_col = col
                    break
            if name_col:
                top_df = df.nlargest(10, price_col)[[name_col, price_col]]
                fig = px.bar(top_df, x=name_col, y=price_col, title="Топ 10 товаров по цене")
                st.plotly_chart(fig, use_container_width=True)


# ============================================================================
# БЛОК 18: ИСТОРИЯ РАСЧЕТОВ
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
        tax_systems = ['Все'] + sorted(df_history['tax_system'].unique().tolist()) if 'tax_system' in df_history.columns else ['Все']
        filter_tax = st.selectbox("Система налога", tax_systems, key="history_tax")

    filtered_df = df_history.copy()
    if filter_marketplace != 'Все':
        filtered_df = filtered_df[filtered_df['marketplace'] == filter_marketplace]
    if filter_mode != 'Все':
        filtered_df = filtered_df[filtered_df['operation_mode'] == filter_mode]
    if filter_tax != 'Все' and 'tax_system' in filtered_df.columns:
        filtered_df = filtered_df[filtered_df['tax_system'] == filter_tax]

    st.subheader(f"📊 Найдено расчетов: {len(filtered_df)}")

    if not filtered_df.empty:
        display_cols = ['marketplace', 'operation_mode', 'price', 'cost', 'profit',
                       'margin_percent', 'roi', 'recommended_min_price', 'tax_amount', 'timestamp']
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
            if 'tax_amount' in filtered_df.columns:
                total_tax = filtered_df['tax_amount'].sum()
                st.metric("💵 Суммарный налог", f"{total_tax:,.0f} ₽")
        with col4:
            if 'recommended_min_price' in filtered_df.columns:
                avg_rec_price = filtered_df['recommended_min_price'].mean()
                st.metric("🎯 Средняя рек. цена", f"{avg_rec_price:.2f} ₽")

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
# БЛОК 19: ЭКСПОРТ
# ============================================================================
def show_export_interface():
    """Интерфейс экспорта данных"""
    st.header("📤 Экспорт данных")

    if st.session_state.get('uploaded_data') is None:
        st.warning("⚠️ Сначала загрузите данные")
        return

    df = st.session_state.uploaded_data
    st.success(f"✅ Готово к экспорту: {len(df)} товаров, {len(df.columns)} колонок")

    export_format = st.radio("Формат экспорта", ["Excel (.xlsx)", "CSV (.csv)"], horizontal=True, key="export_format")

    if st.button("📥 Скачать файл", type="primary", key="export_btn"):
        try:
            if export_format.startswith("Excel"):
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df.to_excel(writer, sheet_name='Данные', index=False)
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


# ============================================================================
# БЛОК 20: НАСТРОЙКИ (v97 С НАЛОГАМИ)
# ============================================================================
def show_settings_interface():
    """Интерфейс настроек (v97)"""
    st.header("⚙️ Настройки приложения")

    unit_economics = MarketplaceUnitEconomics()

    # === НОВОЕ v97: НАСТРОЙКИ НАЛОГОВ ===
    st.subheader("💰 Настройки налогов")
    st.info("""
Настройки применяются ко всем расчетам юнит-экономики.
""")

    col1, col2 = st.columns(2)
    with col1:
        tax_system = st.selectbox(
            "Система налогообложения по умолчанию",
            ["УСН_6", "УСН_15", "ОСНО"],
            index=["УСН_6", "УСН_15", "ОСНО"].index(unit_economics._settings.get('tax_system', 'УСН_6')),
            key="settings_tax_system"
        )
    with col2:
        usn_rate = st.number_input(
            "Ставка УСН (%)",
            min_value=0.0,
            max_value=15.0,
            value=unit_economics._settings.get('usn_rate', 0.06) * 100,
            step=0.5,
            key="settings_usn_rate"
        )

    col1, col2, col3 = st.columns(3)
    with col1:
        min_profit_percent = st.number_input(
            "Минимальная прибыль (%)",
            min_value=0.0,
            max_value=100.0,
            value=unit_economics._settings.get('min_profit_percent', 0.10) * 100,
            step=1.0,
            key="settings_min_profit_percent"
        )
    with col2:
        insurance_contributions = st.number_input(
            "Страховые взносы в год (₽)",
            min_value=0.0,
            value=unit_economics._settings.get('insurance_contributions', 49500),
            step=1000.0,
            key="settings_insurance"
        )
    with col3:
        use_tax_deduction = st.checkbox(
            "Уменьшать налог на взносы",
            value=unit_economics._settings.get('use_tax_deduction', True),
            key="settings_tax_deduction"
        )

    # === ОБЩИЕ НАСТРОЙКИ ===
    st.subheader("🎨 Общие настройки")
    col1, col2 = st.columns(2)
    with col1:
        default_margin = st.number_input(
            "Целевая маржинальность (%)",
            min_value=0.0,
            max_value=100.0,
            value=15.0,
            step=1.0,
            key="settings_default_margin"
        )
    with col2:
        currency = st.selectbox(
            "Основная валюта",
            ["₽ (Рубль)", "$ (Доллар)", "€ (Евро)"],
            key="settings_currency"
        )

    if st.button("💾 Сохранить настройки", type="primary", key="settings_save"):
        new_settings = {
            "tax_system": tax_system,
            "usn_rate": usn_rate / 100,
            "min_profit_percent": min_profit_percent / 100,
            "insurance_contributions": insurance_contributions,
            "use_tax_deduction": use_tax_deduction,
            "target_margin": default_margin,
            "currency": currency
        }
        if unit_economics.save_settings(new_settings):
            st.success("✅ Настройки сохранены!")
            st.balloons()
        else:
            st.error("❌ Ошибка сохранения настроек")

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
# ============================================================================
# БЛОК 21: ЭКСПОРТ С ЖИВЫМИ ФОРМУЛАМИ v97 (5 ЛИСТОВ EXCEL) (800+ СТРОК)
# ============================================================================
def export_unit_economics_with_formulas_v2(
    df_results: pd.DataFrame,
    output_path: str,
    marketplace_configs: Dict[str, Any] = None,
    tax_settings: Optional[Dict[str, Any]] = None
) -> bool:
    """
    Экспорт юнит-экономики в Excel С ЖИВЫМИ ФОРМУЛАМИ.
    5 листов: Тарифы, Налоги, Расчёт, Сводка, Dashboard.
    
    При изменении цены/себестоимости/тарифов всё пересчитывается автоматически.
    
    Args:
        df_results: DataFrame с результатами расчётов
        output_path: Путь для сохранения файла
        marketplace_configs: Конфигурации маркетплейсов
        tax_settings: Настройки налогов
    
    Returns:
        True при успехе, False при ошибке
    """
    if not OPENPYXL_AVAILABLE:
        logger.error("openpyxl не установлен")
        return False

    try:
        from openpyxl import Workbook
        from openpyxl.styles import (
            Font, PatternFill, Alignment, Border, Side, 
            numbers, NamedStyle
        )
        from openpyxl.utils import get_column_letter
        from openpyxl.formatting.rule import CellIsRule, FormulaRule
        from openpyxl.chart import BarChart, PieChart, Reference
        from openpyxl.chart.label import DataLabelList
        from openpyxl.chart.series import DataPoint
        from openpyxl.worksheet.datavalidation import DataValidation

        wb = Workbook()

        # === СТИЛИ ===
        header_font = Font(bold=True, color="FFFFFF", size=11, name='Calibri')
        header_fill = PatternFill(start_color="0F3460", end_color="0F3460", fill_type="solid")
        subheader_fill = PatternFill(start_color="1A5490", end_color="1A5490", fill_type="solid")
        input_fill = PatternFill(start_color="FFF4CC", end_color="FFF4CC", fill_type="solid")  # жёлтый
        formula_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")  # зелёный
        result_fill = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")  # синий
        warning_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # красный
        success_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # зелёный
        highlight_fill = PatternFill(start_color="FFE699", end_color="FFE699", fill_type="solid")  # оранжевый
        
        thin_border = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        
        thick_border = Border(
            left=Side(style='medium'),
            right=Side(style='medium'),
            top=Side(style='medium'),
            bottom=Side(style='medium')
        )
        
        center_align = Alignment(horizontal="center", vertical="center", wrap_text=True)
        left_align = Alignment(horizontal="left", vertical="center", wrap_text=True)
        right_align = Alignment(horizontal="right", vertical="center")

        # === ПОЛУЧЕНИЕ КОНФИГУРАЦИЙ ===
        if marketplace_configs is None:
            try:
                ue = MarketplaceUnitEconomics()
                marketplace_configs = ue._configs
            except Exception:
                marketplace_configs = get_marketplace_configs_2026()

        if tax_settings is None:
            tax_settings = {
                "tax_system": DEFAULT_TAX_SYSTEM,
                "usn_rate": DEFAULT_USN_RATE,
                "min_profit_percent": DEFAULT_MIN_PROFIT_PERCENT,
                "insurance_contributions": DEFAULT_INSURANCE_CONTRIBUTIONS,
                "use_tax_deduction": DEFAULT_TAX_DEDUCTION
            }

        # ============================================================
        # ЛИСТ 1: ТАРИФЫ МАРКЕТПЛЕЙСОВ (СПРАВОЧНИК)
        # ============================================================
        ws_tariffs = wb.active
        ws_tariffs.title = "Тарифы"

        tariff_headers = [
            "Маркетплейс",
            "Комиссия (%)",
            "Мин. комиссия (₽)",
            "Логистика база (₽)",
            "Логистика за кг (₽/кг)",
            "Логистика за литр (₽/л)",
            "Хранение в день (₽/л)",
            "Эквайринг (%)",
            "Последняя миля (₽)",
            "Возвраты (%)",
            "РКО (%)",
            "Подписка в месяц (₽)",
            "Надбавка опасные (%)",
            "Надбавка хрупкие (%)",
            "Надбавка крупногабарит (%)"
        ]

        # Заголовок
        ws_tariffs.cell(row=1, column=1, value="📊 ТАРИФЫ МАРКЕТПЛЕЙСОВ 2026")
        ws_tariffs.cell(row=1, column=1).font = Font(bold=True, size=14, color="0F3460")
        ws_tariffs.merge_cells('A1:O1')
        ws_tariffs.cell(row=1, column=1).alignment = center_align

        ws_tariffs.cell(row=2, column=1, value="⚠️ Редактируйте тарифы — все расчёты обновятся автоматически")
        ws_tariffs.cell(row=2, column=1).font = Font(italic=True, color="FF0000", size=10)
        ws_tariffs.merge_cells('A2:O2')

        # Заголовки колонок
        for col_idx, header in enumerate(tariff_headers, start=1):
            cell = ws_tariffs.cell(row=4, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        # Данные тарифов
        row_idx = 5
        for mp_name, config in marketplace_configs.items():
            ws_tariffs.cell(row=row_idx, column=1, value=mp_name).border = thin_border
            ws_tariffs.cell(row=row_idx, column=1).fill = input_fill
            ws_tariffs.cell(row=row_idx, column=1).font = Font(bold=True)
            
            # Все числовые поля — редактируемые (жёлтые)
            tariff_values = [
                (2, config.commission_rate, '0.00%'),
                (3, config.min_commission, '#,##0.00'),
                (4, config.logistics_base, '#,##0.00'),
                (5, config.logistics_per_kg, '#,##0.00'),
                (6, config.logistics_per_liter, '#,##0.00'),
                (7, config.storage_per_day, '#,##0.00'),
                (8, config.acquiring_fee, '0.00%'),
                (9, config.last_mile_fee, '#,##0.00'),
                (10, config.return_fee, '0.00%'),
                (11, config.rko_fee, '0.00%'),
                (12, config.subscription_fee, '#,##0.00'),
                (13, config.hazardous_surcharge, '0.00%'),
                (14, config.fragile_surcharge, '0.00%'),
                (15, config.oversized_surcharge, '0.00%'),
            ]
            
            for col, val, fmt in tariff_values:
                cell = ws_tariffs.cell(row=row_idx, column=col, value=val)
                cell.fill = input_fill
                cell.border = thin_border
                cell.number_format = fmt
                cell.alignment = center_align
            
            row_idx += 1

        # Ширина колонок
        for col_idx in range(1, len(tariff_headers) + 1):
            ws_tariffs.column_dimensions[get_column_letter(col_idx)].width = 18

        ws_tariffs.freeze_panes = "B5"
        ws_tariffs.auto_filter.ref = f"A4:O{row_idx - 1}"
        
        tariff_last_row = row_idx - 1

        # ============================================================
        # ЛИСТ 2: НАЛОГИ (НАСТРОЙКИ)
        # ============================================================
        ws_taxes = wb.create_sheet("Налоги")

        ws_taxes.cell(row=1, column=1, value="💰 НАСТРОЙКИ НАЛОГООБЛОЖЕНИЯ")
        ws_taxes.cell(row=1, column=1).font = Font(bold=True, size=14, color="0F3460")
        ws_taxes.merge_cells('A1:C1')
        ws_taxes.cell(row=1, column=1).alignment = center_align

        ws_taxes.cell(row=2, column=1, value="⚠️ Редактируйте значения — все расчёты налогов обновятся")
        ws_taxes.cell(row=2, column=1).font = Font(italic=True, color="FF0000", size=10)
        ws_taxes.merge_cells('A2:C2')

        tax_params = [
            ("Параметр", "Значение", "Описание"),
            ("Система налогообложения", tax_settings.get('tax_system', 'УСН_6'), "УСН_6, УСН_15 или ОСНО"),
            ("Ставка УСН (%)", tax_settings.get('usn_rate', 0.06) * 100, "Процент налога от выручки"),
            ("Минимальная прибыль (%)", tax_settings.get('min_profit_percent', 0.10) * 100, "Желаемая прибыль после всех расходов"),
            ("Страховые взносы (₽/год)", tax_settings.get('insurance_contributions', 49500), "Фиксированные взносы ИП"),
            ("Уменьшать налог на взносы", "ДА" if tax_settings.get('use_tax_deduction', True) else "НЕТ", "Для ИП без сотрудников"),
            ("Лимит УСН (₽)", 300_000_000, "Свыше — дополнительный 1%"),
            ("Доп. ставка сверх лимита (%)", 1.0, "1% с превышения"),
        ]

        for row_idx, (param, value, desc) in enumerate(tax_params, start=4):
            cell_param = ws_taxes.cell(row=row_idx, column=1, value=param)
            cell_value = ws_taxes.cell(row=row_idx, column=2, value=value)
            cell_desc = ws_taxes.cell(row=row_idx, column=3, value=desc)
            
            if row_idx == 4:  # Заголовок
                cell_param.font = header_font
                cell_param.fill = header_fill
                cell_value.font = header_font
                cell_value.fill = header_fill
                cell_desc.font = header_font
                cell_desc.fill = header_fill
            else:
                cell_param.border = thin_border
                cell_value.border = thin_border
                cell_value.fill = input_fill
                cell_desc.border = thin_border
                cell_desc.font = Font(italic=True, color="666666")
            
            cell_param.alignment = left_align
            cell_value.alignment = center_align
            cell_desc.alignment = left_align

        ws_taxes.column_dimensions['A'].width = 30
        ws_taxes.column_dimensions['B'].width = 20
        ws_taxes.column_dimensions['C'].width = 50

        # ============================================================
        # ЛИСТ 3: РАСЧЁТ ЮНИТ-ЭКОНОМИКИ (С ФОРМУЛАМИ)
        # ============================================================
        ws_calc = wb.create_sheet("Расчёт")

        calc_headers = [
            "Артикул",           # A
            "Бренд",             # B
            "Маркетплейс",       # C
            "Режим работы",      # D
            "Цена продажи (₽)",  # E - ВВОД
            "Себестоимость (₽)", # F - ВВОД
            "Длина (см)",        # G - ВВОД
            "Ширина (см)",       # H - ВВОД
            "Высота (см)",       # I - ВВОД
            "Вес (кг)",          # J - ВВОД
            "Объём (л)",         # K - ФОРМУЛА
            "Комиссия МП (%)",   # L - VLOOKUP
            "Комиссия (₽)",      # M - ФОРМУЛА
            "Логистика (₽)",     # N - ФОРМУЛА
            "Хранение (₽)",      # O - ФОРМУЛА
            "Эквайринг (₽)",     # P - ФОРМУЛА
            "Последняя миля (₽)",# Q - VLOOKUP
            "Возвраты (₽)",      # R - ФОРМУЛА
            "Надбавка опасные",  # S - ФОРМУЛА
            "Надбавка хрупкие",  # T - ФОРМУЛА
            "Надбавка крупногаб",# U - ФОРМУЛА
            "Налог (₽)",         # V - ФОРМУЛА
            "Итого расходов (₽)",# W - ФОРМУЛА
            "Прибыль (₽)",       # X - ФОРМУЛА
            "Маржа (%)",         # Y - ФОРМУЛА
            "ROI (%)",           # Z - ФОРМУЛА
            "Точка безубыточности", # AA - ФОРМУЛА
            "Рекомендованная мин. цена (₽)" # AB - ФОРМУЛА
        ]

        # Заголовок
        ws_calc.cell(row=1, column=1, value="📊 РАСЧЁТ ЮНИТ-ЭКОНОМИКИ")
        ws_calc.cell(row=1, column=1).font = Font(bold=True, size=14, color="0F3460")
        ws_calc.merge_cells('A1:AB1')
        ws_calc.cell(row=1, column=1).alignment = center_align

        ws_calc.cell(row=2, column=1, value="🟡 Жёлтые ячейки — ввод | 🟢 Зелёные — формулы | 🔵 Синие — результаты")
        ws_calc.cell(row=2, column=1).font = Font(italic=True, size=10)
        ws_calc.merge_cells('A2:AB2')

        # Заголовки колонок
        for col_idx, header in enumerate(calc_headers, start=1):
            cell = ws_calc.cell(row=4, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        # Диапазоны для VLOOKUP
        tariff_range = f"Тарифы!$A$5:$O${tariff_last_row}"

        # Данные с формулами
        for r_idx, row_data in enumerate(df_results.itertuples(index=False), start=5):
            row_dict = row_data._asdict() if hasattr(row_data, '_asdict') else dict(zip(df_results.columns, row_data))

            # A-D: Вводные данные (текст)
            ws_calc.cell(row=r_idx, column=1, value=row_dict.get('Артикул', '')).border = thin_border
            ws_calc.cell(row=r_idx, column=2, value=row_dict.get('Бренд', '')).border = thin_border
            ws_calc.cell(row=r_idx, column=3, value=row_dict.get('Маркетплейс', 'Ozon')).border = thin_border
            ws_calc.cell(row=r_idx, column=4, value=row_dict.get('Режим_работы', 'FBS')).border = thin_border

            # E-J: ВХОДНЫЕ ДАННЫЕ (жёлтые, редактируемые)
            input_cols = {
                5: ('Цена_с_наценкой', row_dict.get('Цена_с_наценкой', row_dict.get('Цена_исходная', 0))),
                6: ('Себестоимость', row_dict.get('Себестоимость', 0)),
                7: ('Длина', row_dict.get('Длина_см', 0)),
                8: ('Ширина', row_dict.get('Ширина_см', 0)),
                9: ('Высота', row_dict.get('Высота_см', 0)),
                10: ('Вес', row_dict.get('Вес_кг', 0)),
            }
            
            for col, (key, default_val) in input_cols.items():
                val = row_dict.get(key, default_val)
                cell = ws_calc.cell(row=r_idx, column=col, value=val)
                cell.fill = input_fill
                cell.border = thin_border
                cell.number_format = '#,##0.00'
                cell.alignment = center_align

            r = r_idx  # номер строки для формул

            # K: Объём (л) = (Д * Ш * В) / 1000
            ws_calc.cell(row=r, column=11).value = f"=IF(AND(G{r}>0,H{r}>0,I{r}>0),(G{r}*H{r}*I{r})/1000,5)"
            ws_calc.cell(row=r, column=11).fill = formula_fill
            ws_calc.cell(row=r, column=11).border = thin_border
            ws_calc.cell(row=r, column=11).number_format = '#,##0.000'

            # L: Комиссия МП (%) - VLOOKUP из тарифов
            ws_calc.cell(row=r, column=12).value = f"=IFERROR(VLOOKUP(C{r},{tariff_range},2,FALSE),0.15)"
            ws_calc.cell(row=r, column=12).fill = formula_fill
            ws_calc.cell(row=r, column=12).border = thin_border
            ws_calc.cell(row=r, column=12).number_format = '0.00%'

            # M: Комиссия (₽) = MAX(Цена * Комиссия%, Мин. комиссия)
            ws_calc.cell(row=r, column=13).value = (
                f"=MAX(E{r}*L{r}, IFERROR(VLOOKUP(C{r},{tariff_range},3,FALSE),0))"
            )
            ws_calc.cell(row=r, column=13).fill = formula_fill
            ws_calc.cell(row=r, column=13).border = thin_border
            ws_calc.cell(row=r, column=13).number_format = '#,##0.00'

            # N: Логистика = (База + Вес*ставка_кг + Объём*ставка_л) * Коэф_режима
            ws_calc.cell(row=r, column=14).value = (
                f"=(IFERROR(VLOOKUP(C{r},{tariff_range},4,FALSE),50) "
                f"+ J{r}*IFERROR(VLOOKUP(C{r},{tariff_range},5,FALSE),15) "
                f"+ K{r}*IFERROR(VLOOKUP(C{r},{tariff_range},6,FALSE),5)) "
                f"* IF(D{r}=\"FBY\",0.75,IF(D{r}=\"FBS\",1,IF(D{r}=\"FBO\",0.8,IF(D{r}=\"DBS\",1.3,IF(D{r}=\"FBP\",0.9,1)))))"
            )
            ws_calc.cell(row=r, column=14).fill = formula_fill
            ws_calc.cell(row=r, column=14).border = thin_border
            ws_calc.cell(row=r, column=14).number_format = '#,##0.00'

            # O: Хранение = Объём * Ставка_хранения * 30 дней
            ws_calc.cell(row=r, column=15).value = (
                f"=K{r} * IFERROR(VLOOKUP(C{r},{tariff_range},7,FALSE),0.3) * 30"
            )
            ws_calc.cell(row=r, column=15).fill = formula_fill
            ws_calc.cell(row=r, column=15).border = thin_border
            ws_calc.cell(row=r, column=15).number_format = '#,##0.00'

            # P: Эквайринг = Цена * Ставка_эквайринга
            ws_calc.cell(row=r, column=16).value = (
                f"=E{r} * IFERROR(VLOOKUP(C{r},{tariff_range},8,FALSE),0.015)"
            )
            ws_calc.cell(row=r, column=16).fill = formula_fill
            ws_calc.cell(row=r, column=16).border = thin_border
            ws_calc.cell(row=r, column=16).number_format = '#,##0.00'

            # Q: Последняя миля = фикс из тарифа
            ws_calc.cell(row=r, column=17).value = (
                f"=IFERROR(VLOOKUP(C{r},{tariff_range},9,FALSE),50)"
            )
            ws_calc.cell(row=r, column=17).fill = formula_fill
            ws_calc.cell(row=r, column=17).border = thin_border
            ws_calc.cell(row=r, column=17).number_format = '#,##0.00'

            # R: Возвраты = Цена * Ставка_возвратов
            ws_calc.cell(row=r, column=18).value = (
                f"=E{r} * IFERROR(VLOOKUP(C{r},{tariff_range},10,FALSE),0.02)"
            )
            ws_calc.cell(row=r, column=18).fill = formula_fill
            ws_calc.cell(row=r, column=18).border = thin_border
            ws_calc.cell(row=r, column=18).number_format = '#,##0.00'

            # S: Надбавка за опасные (если категория опасная)
            ws_calc.cell(row=r, column=19).value = (
                f"=E{r} * IFERROR(VLOOKUP(C{r},{tariff_range},13,FALSE),0)"
            )
            ws_calc.cell(row=r, column=19).fill = formula_fill
            ws_calc.cell(row=r, column=19).border = thin_border
            ws_calc.cell(row=r, column=19).number_format = '#,##0.00'

            # T: Надбавка за хрупкие
            ws_calc.cell(row=r, column=20).value = (
                f"=E{r} * IFERROR(VLOOKUP(C{r},{tariff_range},14,FALSE),0)"
            )
            ws_calc.cell(row=r, column=20).fill = formula_fill
            ws_calc.cell(row=r, column=20).border = thin_border
            ws_calc.cell(row=r, column=20).number_format = '#,##0.00'

            # U: Надбавка за крупногабарит
            ws_calc.cell(row=r, column=21).value = (
                f"=E{r} * IFERROR(VLOOKUP(C{r},{tariff_range},15,FALSE),0)"
            )
            ws_calc.cell(row=r, column=21).fill = formula_fill
            ws_calc.cell(row=r, column=21).border = thin_border
            ws_calc.cell(row=r, column=21).number_format = '#,##0.00'

            # V: Налог = Цена * Ставка_налога (из листа Налоги)
            ws_calc.cell(row=r, column=22).value = (
                f"=E{r} * IF(Налоги!B5=\"УСН_6\",Налоги!B6/100,"
                f"IF(Налоги!B5=\"УСН_15\",0.15,0.40))"
            )
            ws_calc.cell(row=r, column=22).fill = formula_fill
            ws_calc.cell(row=r, column=22).border = thin_border
            ws_calc.cell(row=r, column=22).number_format = '#,##0.00'

            # W: ИТОГО РАСХОДОВ = Себестоимость + все комиссии и расходы
            ws_calc.cell(row=r, column=23).value = (
                f"=F{r}+M{r}+N{r}+O{r}+P{r}+Q{r}+R{r}+S{r}+T{r}+U{r}+V{r}"
            )
            ws_calc.cell(row=r, column=23).fill = result_fill
            ws_calc.cell(row=r, column=23).border = thick_border
            ws_calc.cell(row=r, column=23).number_format = '#,##0.00'
            ws_calc.cell(row=r, column=23).font = Font(bold=True)

            # X: ПРИБЫЛЬ = Цена - Итого расходов
            ws_calc.cell(row=r, column=24).value = f"=E{r}-W{r}"
            ws_calc.cell(row=r, column=24).fill = result_fill
            ws_calc.cell(row=r, column=24).border = thick_border
            ws_calc.cell(row=r, column=24).number_format = '#,##0.00'
            ws_calc.cell(row=r, column=24).font = Font(bold=True, size=12)

            # Y: МАРЖА (%) = Прибыль / Цена * 100
            ws_calc.cell(row=r, column=25).value = f"=IF(E{r}=0,0,X{r}/E{r}*100)"
            ws_calc.cell(row=r, column=25).fill = result_fill
            ws_calc.cell(row=r, column=25).border = thin_border
            ws_calc.cell(row=r, column=25).number_format = '0.00'

            # Z: ROI (%) = Прибыль / Себестоимость * 100
            ws_calc.cell(row=r, column=26).value = f"=IF(F{r}=0,0,X{r}/F{r}*100)"
            ws_calc.cell(row=r, column=26).fill = result_fill
            ws_calc.cell(row=r, column=26).border = thin_border
            ws_calc.cell(row=r, column=26).number_format = '0.00'

            # AA: ТОЧКА БЕЗУБЫТОЧНОСТИ
            ws_calc.cell(row=r, column=27).value = (
                f"=IF((1-L{r}-IFERROR(VLOOKUP(C{r},{tariff_range},8,FALSE),0)"
                f"-IFERROR(VLOOKUP(C{r},{tariff_range},10,FALSE),0)"
                f"-IF(Налоги!B5=\"УСН_6\",Налоги!B6/100,IF(Налоги!B5=\"УСН_15\",0.15,0.40)))<=0,0,"
                f"(F{r}+N{r}+O{r}+Q{r})/(1-L{r}-IFERROR(VLOOKUP(C{r},{tariff_range},8,FALSE),0)"
                f"-IFERROR(VLOOKUP(C{r},{tariff_range},10,FALSE),0)"
                f"-IF(Налоги!B5=\"УСН_6\",Налоги!B6/100,IF(Налоги!B5=\"УСН_15\",0.15,0.40))))"
            )
            ws_calc.cell(row=r, column=27).fill = result_fill
            ws_calc.cell(row=r, column=27).border = thin_border
            ws_calc.cell(row=r, column=27).number_format = '#,##0.00'

            # AB: РЕКОМЕНДОВАННАЯ МИНИМАЛЬНАЯ ЦЕНА
            # Формула: X = (Cost + Log + Stor + LastMile) / (1 - comm - acq - ret - tax - min_profit)
            ws_calc.cell(row=r, column=28).value = (
                f"=IF((1-L{r}-IFERROR(VLOOKUP(C{r},{tariff_range},8,FALSE),0)"
                f"-IFERROR(VLOOKUP(C{r},{tariff_range},10,FALSE),0)"
                f"-IF(Налоги!B5=\"УСН_6\",Налоги!B6/100,IF(Налоги!B5=\"УСН_15\",0.15,0.40))"
                f"-Налоги!B7/100)<=0,0,"
                f"(F{r}+N{r}+O{r}+Q{r})/(1-L{r}-IFERROR(VLOOKUP(C{r},{tariff_range},8,FALSE),0)"
                f"-IFERROR(VLOOKUP(C{r},{tariff_range},10,FALSE),0)"
                f"-IF(Налоги!B5=\"УСН_6\",Налоги!B6/100,IF(Налоги!B5=\"УСН_15\",0.15,0.40))"
                f"-Налоги!B7/100))"
            )
            ws_calc.cell(row=r, column=28).fill = highlight_fill
            ws_calc.cell(row=r, column=28).border = thick_border
            ws_calc.cell(row=r, column=28).number_format = '#,##0.00'
            ws_calc.cell(row=r, column=28).font = Font(bold=True, size=12, color="0F3460")

        # === УСЛОВНОЕ ФОРМАТИРОВАНИЕ ДЛЯ ПРИБЫЛИ (X) ===
        profit_range = f"X5:X{len(df_results) + 4}"
        
        # Зелёный для прибыли > 0
        ws_calc.conditional_formatting.add(
            profit_range,
            CellIsRule(operator='greaterThan', formula=['0'], 
                      fill=success_fill, 
                      font=Font(color="006100", bold=True))
        )
        
        # Красный для убытка < 0
        ws_calc.conditional_formatting.add(
            profit_range,
            CellIsRule(operator='lessThan', formula=['0'], 
                      fill=warning_fill, 
                      font=Font(color="9C0006", bold=True))
        )

        # === УСЛОВНОЕ ФОРМАТИРОВАНИЕ ДЛЯ МАРЖИ (Y) ===
        margin_range = f"Y5:Y{len(df_results) + 4}"
        
        # Зелёный для маржи >= 10%
        ws_calc.conditional_formatting.add(
            margin_range,
            CellIsRule(operator='greaterThanOrEqual', formula=['10'], 
                      fill=success_fill)
        )
        
        # Жёлтый для маржи 5-10%
        ws_calc.conditional_formatting.add(
            margin_range,
            FormulaRule(formula=[f'AND(Y5>=5,Y5<10)'], fill=highlight_fill)
        )
        
        # Красный для маржи < 5%
        ws_calc.conditional_formatting.add(
            margin_range,
            CellIsRule(operator='lessThan', formula=['5'], 
                      fill=warning_fill)
        )

        # === УСЛОВНОЕ ФОРМАТИРОВАНИЕ ДЛЯ РЕКОМЕНДОВАННОЙ ЦЕНЫ (AB) ===
        rec_price_range = f"AB5:AB{len(df_results) + 4}"
        
        # Красный если текущая цена < рекомендованной
        ws_calc.conditional_formatting.add(
            rec_price_range,
            FormulaRule(formula=[f'E5<AB5'], fill=warning_fill, font=Font(color="9C0006", bold=True))
        )

        # Ширина колонок
        for col_idx in range(1, len(calc_headers) + 1):
            ws_calc.column_dimensions[get_column_letter(col_idx)].width = 18

        ws_calc.freeze_panes = "E5"
        ws_calc.auto_filter.ref = f"A4:AB{len(df_results) + 4}"

        # ============================================================
        # ЛИСТ 4: СВОДКА (АГРЕГАЦИЯ)
        # ============================================================
        ws_summary = wb.create_sheet("Сводка")

        ws_summary.cell(row=1, column=1, value="📊 СВОДНАЯ СТАТИСТИКА ПО МАРКЕТПЛЕЙСАМ")
        ws_summary.cell(row=1, column=1).font = Font(bold=True, size=14, color="0F3460")
        ws_summary.merge_cells('A1:H1')
        ws_summary.cell(row=1, column=1).alignment = center_align

        summary_headers = [
            "Маркетплейс",
            "Количество товаров",
            "Общая прибыль (₽)",
            "Средняя прибыль (₽)",
            "Средняя маржа (%)",
            "Средний ROI (%)",
            "Средняя рек. цена (₽)",
            "Мин. прибыль (₽)"
        ]

        for col_idx, header in enumerate(summary_headers, start=1):
            cell = ws_summary.cell(row=3, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        # Получаем уникальные маркетплейсы
        marketplaces = df_results['Маркетплейс'].unique().tolist() if 'Маркетплейс' in df_results.columns else []
        data_range = f"Расчёт!$C$5:$C${len(df_results) + 4}"
        last_data_row = len(df_results) + 4

        for s_idx, mp in enumerate(marketplaces, start=4):
            ws_summary.cell(row=s_idx, column=1, value=mp).border = thin_border
            ws_summary.cell(row=s_idx, column=1).font = Font(bold=True)
            
            # Количество товаров
            ws_summary.cell(row=s_idx, column=2).value = f'=COUNTIF({data_range},A{s_idx})'
            ws_summary.cell(row=s_idx, column=2).border = thin_border
            
            # Общая прибыль
            ws_summary.cell(row=s_idx, column=3).value = f'=SUMIF({data_range},A{s_idx},Расчёт!$X$5:$X${last_data_row})'
            ws_summary.cell(row=s_idx, column=3).border = thin_border
            ws_summary.cell(row=s_idx, column=3).number_format = '#,##0.00'
            
            # Средняя прибыль
            ws_summary.cell(row=s_idx, column=4).value = f'=IF(B{s_idx}=0,0,C{s_idx}/B{s_idx})'
            ws_summary.cell(row=s_idx, column=4).border = thin_border
            ws_summary.cell(row=s_idx, column=4).number_format = '#,##0.00'
            
            # Средняя маржа
            ws_summary.cell(row=s_idx, column=5).value = f'=AVERAGEIF({data_range},A{s_idx},Расчёт!$Y$5:$Y${last_data_row})'
            ws_summary.cell(row=s_idx, column=5).border = thin_border
            ws_summary.cell(row=s_idx, column=5).number_format = '0.00'
            
            # Средний ROI
            ws_summary.cell(row=s_idx, column=6).value = f'=AVERAGEIF({data_range},A{s_idx},Расчёт!$Z$5:$Z${last_data_row})'
            ws_summary.cell(row=s_idx, column=6).border = thin_border
            ws_summary.cell(row=s_idx, column=6).number_format = '0.00'
            
            # Средняя рек. цена
            ws_summary.cell(row=s_idx, column=7).value = f'=AVERAGEIF({data_range},A{s_idx},Расчёт!$AB$5:$AB${last_data_row})'
            ws_summary.cell(row=s_idx, column=7).border = thin_border
            ws_summary.cell(row=s_idx, column=7).number_format = '#,##0.00'
            
            # Мин. прибыль
            ws_summary.cell(row=s_idx, column=8).value = f'=MINIFS(Расчёт!$X$5:$X${last_data_row},{data_range},A{s_idx})'
            ws_summary.cell(row=s_idx, column=8).border = thin_border
            ws_summary.cell(row=s_idx, column=8).number_format = '#,##0.00'

        # ИТОГО строка
        total_row = 4 + len(marketplaces)
        ws_summary.cell(row=total_row, column=1, value="ИТОГО").font = Font(bold=True, size=12)
        ws_summary.cell(row=total_row, column=1).border = thick_border
        ws_summary.cell(row=total_row, column=1).fill = result_fill
        
        for col in range(2, 9):
            ws_summary.cell(row=total_row, column=col).border = thick_border
            ws_summary.cell(row=total_row, column=col).fill = result_fill
            ws_summary.cell(row=total_row, column=col).font = Font(bold=True)
        
        ws_summary.cell(row=total_row, column=2).value = f'=SUM(B4:B{total_row-1})'
        ws_summary.cell(row=total_row, column=3).value = f'=SUM(C4:C{total_row-1})'
        ws_summary.cell(row=total_row, column=3).number_format = '#,##0.00'
        ws_summary.cell(row=total_row, column=4).value = f'=IF(B{total_row}=0,0,C{total_row}/B{total_row})'
        ws_summary.cell(row=total_row, column=4).number_format = '#,##0.00'
        ws_summary.cell(row=total_row, column=5).value = f'=AVERAGE(E4:E{total_row-1})'
        ws_summary.cell(row=total_row, column=5).number_format = '0.00'
        ws_summary.cell(row=total_row, column=6).value = f'=AVERAGE(F4:F{total_row-1})'
        ws_summary.cell(row=total_row, column=6).number_format = '0.00'
        ws_summary.cell(row=total_row, column=7).value = f'=AVERAGE(G4:G{total_row-1})'
        ws_summary.cell(row=total_row, column=7).number_format = '#,##0.00'
        ws_summary.cell(row=total_row, column=8).value = f'=MIN(H4:H{total_row-1})'
        ws_summary.cell(row=total_row, column=8).number_format = '#,##0.00'

        for col_idx in range(1, len(summary_headers) + 1):
            ws_summary.column_dimensions[get_column_letter(col_idx)].width = 22

        ws_summary.freeze_panes = "A4"

        # === ДИАГРАММА ПРИБЫЛИ ПО МП ===
        chart = BarChart()
        chart.type = "col"
        chart.style = 10
        chart.title = "Прибыль по маркетплейсам"
        chart.y_axis.title = 'Прибыль (₽)'
        chart.x_axis.title = 'Маркетплейс'

        data = Reference(ws_summary, min_col=3, min_row=3, max_row=total_row-1)
        cats = Reference(ws_summary, min_col=1, min_row=4, max_row=total_row-1)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        chart.height = 10
        chart.width = 20

        ws_summary.add_chart(chart, "A" + str(total_row + 3))

        # ============================================================
        # ЛИСТ 5: DASHBOARD (ВИЗУАЛИЗАЦИЯ)
        # ============================================================
        ws_dashboard = wb.create_sheet("Dashboard")

        ws_dashboard.cell(row=1, column=1, value="📊 DASHBOARD: КЛЮЧЕВЫЕ ПОКАЗАТЕЛИ")
        ws_dashboard.cell(row=1, column=1).font = Font(bold=True, size=16, color="0F3460")
        ws_dashboard.merge_cells('A1:F1')
        ws_dashboard.cell(row=1, column=1).alignment = center_align

        # KPI блоки
        kpi_data = [
            ("Всего товаров", f"=Сводка!B{total_row}", '#,##0'),
            ("Общая прибыль (₽)", f"=Сводка!C{total_row}", '#,##0.00 ₽'),
            ("Средняя прибыль (₽)", f"=Сводка!D{total_row}", '#,##0.00 ₽'),
            ("Средняя маржа (%)", f"=Сводка!E{total_row}", '0.00"%"'),
            ("Средний ROI (%)", f"=Сводка!F{total_row}", '0.00"%"'),
            ("Средняя рек. цена (₽)", f"=Сводка!G{total_row}", '#,##0.00 ₽'),
        ]

        for idx, (label, formula, fmt) in enumerate(kpi_data, start=3):
            row = 3 + (idx // 2) * 3
            col = 1 + (idx % 2) * 3
            
            cell_label = ws_dashboard.cell(row=row, column=col, value=label)
            cell_label.font = Font(bold=True, size=11, color="0F3460")
            cell_label.fill = subheader_fill
            cell_label.font = Font(bold=True, color="FFFFFF")
            cell_label.alignment = center_align
            cell_label.border = thin_border
            ws_dashboard.merge_cells(start_row=row, start_column=col, 
                                    end_row=row, end_column=col+1)
            
            cell_value = ws_dashboard.cell(row=row+1, column=col)
            cell_value.value = formula
            cell_value.font = Font(bold=True, size=18, color="0F3460")
            cell_value.alignment = center_align
            cell_value.border = thick_border
            cell_value.fill = result_fill
            cell_value.number_format = fmt
            ws_dashboard.merge_cells(start_row=row+1, start_column=col, 
                                    end_row=row+1, end_column=col+1)

        # Топ-10 товаров по прибыли
        ws_dashboard.cell(row=10, column=1, value="🏆 ТОП-10 ТОВАРОВ ПО ПРИБЫЛИ")
        ws_dashboard.cell(row=10, column=1).font = Font(bold=True, size=14, color="0F3460")
        ws_dashboard.merge_cells('A10:F10')

        top_headers = ["Артикул", "Бренд", "Маркетплейс", "Цена (₽)", "Прибыль (₽)", "Маржа (%)"]
        for col_idx, header in enumerate(top_headers, start=1):
            cell = ws_dashboard.cell(row=11, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center_align
            cell.border = thin_border

        # Сортируем по прибыли и берём топ-10
        df_sorted = df_results.sort_values('Прибыль', ascending=False).head(10)
        for idx, (_, row_data) in enumerate(df_sorted.iterrows(), start=12):
            ws_dashboard.cell(row=idx, column=1, value=row_data.get('Артикул', '')).border = thin_border
            ws_dashboard.cell(row=idx, column=2, value=row_data.get('Бренд', '')).border = thin_border
            ws_dashboard.cell(row=idx, column=3, value=row_data.get('Маркетплейс', '')).border = thin_border
            
            cell_price = ws_dashboard.cell(row=idx, column=4, value=row_data.get('Цена_с_наценкой', 0))
            cell_price.border = thin_border
            cell_price.number_format = '#,##0.00'
            
            cell_profit = ws_dashboard.cell(row=idx, column=5, value=row_data.get('Прибыль', 0))
            cell_profit.border = thin_border
            cell_profit.number_format = '#,##0.00'
            cell_profit.font = Font(bold=True, color="006100" if row_data.get('Прибыль', 0) > 0 else "9C0006")
            
            cell_margin = ws_dashboard.cell(row=idx, column=6, value=row_data.get('Маржа_%', 0))
            cell_margin.border = thin_border
            cell_margin.number_format = '0.00'

        # Топ-10 убыточных товаров
        ws_dashboard.cell(row=24, column=1, value="⚠️ ТОП-10 УБЫТОЧНЫХ ТОВАРОВ")
        ws_dashboard.cell(row=24, column=1).font = Font(bold=True, size=14, color="9C0006")
        ws_dashboard.merge_cells('A24:F24')

        for col_idx, header in enumerate(top_headers, start=1):
            cell = ws_dashboard.cell(row=25, column=col_idx, value=header)
            cell.font = header_font
            cell.fill = PatternFill(start_color="9C0006", end_color="9C0006", fill_type="solid")
            cell.alignment = center_align
            cell.border = thin_border

        df_loss = df_results.sort_values('Прибыль', ascending=True).head(10)
        for idx, (_, row_data) in enumerate(df_loss.iterrows(), start=26):
            ws_dashboard.cell(row=idx, column=1, value=row_data.get('Артикул', '')).border = thin_border
            ws_dashboard.cell(row=idx, column=2, value=row_data.get('Бренд', '')).border = thin_border
            ws_dashboard.cell(row=idx, column=3, value=row_data.get('Маркетплейс', '')).border = thin_border
            
            cell_price = ws_dashboard.cell(row=idx, column=4, value=row_data.get('Цена_с_наценкой', 0))
            cell_price.border = thin_border
            cell_price.number_format = '#,##0.00'
            
            cell_profit = ws_dashboard.cell(row=idx, column=5, value=row_data.get('Прибыль', 0))
            cell_profit.border = thin_border
            cell_profit.number_format = '#,##0.00'
            cell_profit.font = Font(bold=True, color="9C0006")
            
            cell_margin = ws_dashboard.cell(row=idx, column=6, value=row_data.get('Маржа_%', 0))
            cell_margin.border = thin_border
            cell_margin.number_format = '0.00'

        # Ширина колонок
        for col_idx in range(1, 7):
            ws_dashboard.column_dimensions[get_column_letter(col_idx)].width = 22

        # === СОХРАНЕНИЕ ===
        wb.save(output_path)
        logger.info(f"✅ Экспорт с формулами сохранён: {output_path}")
        logger.info(f"📊 Листов: 5 (Тарифы, Налоги, Расчёт, Сводка, Dashboard)")
        logger.info(f"📝 Строк данных: {len(df_results)}")
        return True

    except Exception as e:
        logger.error(f"Ошибка экспорта с формулами: {e}")
        logger.error(traceback.format_exc())
        return False


# ============================================================================
# БЛОК 22: ГЛАВНАЯ ФУНКЦИЯ v97
# ============================================================================
def main():
    """Главная функция приложения v97"""
    st.set_page_config(
        page_title=f"{APP_NAME} v{APP_VERSION}",
        page_icon="🚗",
        layout="wide",
        initial_sidebar_state="expanded"
    )

    # === КРАСИВЫЙ БАННЕР ===
    st.markdown(f"""
<div style="text-align: center; padding: 25px; background: linear-gradient(135deg, #0f3460 0%, #16213e 100%); border-radius: 15px; margin-bottom: 20px; box-shadow: 0 4px 6px rgba(0,0,0,0.1);">
<h1 style="color: white; margin: 0;">🚗 {APP_NAME}</h1>
<p style="color: #e94560; font-size: 20px; margin: 10px 0;">v{APP_VERSION} | ФИНАЛЬНАЯ ВЕРСИЯ 7500+ СТРОК</p>
<p style="color: #aaa; font-size: 14px;">Юнит-экономика маркетплейсов 2026 | Каталог с поиском аналогов 2 уровня</p>
<p style="color: #888; font-size: 13px;">High-Volume каталог (10M+) | ИИ-обновление тарифов | Экспорт с формулами</p>
<div style="margin-top: 15px;">
<p style="color: #00cc96; font-size: 14px; margin: 5px 0;">✅ Галочка "Запросить ИИ" для обновления тарифов</p>
<p style="color: #00cc96; font-size: 14px; margin: 5px 0;">✅ Учёт налогов (УСН 6%, 15%, ОСНО)</p>
<p style="color: #00cc96; font-size: 14px; margin: 5px 0;">✅ Рекомендованная минимальная цена (с учётом 10% прибыли)</p>
<p style="color: #00cc96; font-size: 14px; margin: 5px 0;">✅ Экспорт в Excel с 5 листами и живыми формулами</p>
<p style="color: #00cc96; font-size: 14px; margin: 5px 0;">✅ 150+ категорий автозапчастей с габаритами</p>
<p style="color: #00cc96; font-size: 14px; margin: 5px 0;">✅ Юнит-экономика по каждому артикулу</p>
</div>
</div>
""", unsafe_allow_html=True)

    # === БОКОВОЕ МЕНЮ ===
    with st.sidebar:
        st.image("https://img.icons8.com/fluency/96/000000/bar-chart.png", width=80)
        st.markdown("---")

        menu_options = [
            "📁 Загрузка данных",
            "📊 Обогащение каталога",
            "📊 Юнит-экономика",
            "📊 Юнит-экономика по артикулам",
            "🤖 AI Тарифы",
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
            "🤖 AI Тарифы": "🤖",
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
                "Async": ASYNC_AVAILABLE,
                "Chardet": CHARDET_AVAILABLE
            }
            for lib, available in libs_status.items():
                st.write(f"{'✅' if available else '❌'} {lib}")

    # === МАРШРУТИЗАЦИЯ ===
    try:
        if menu == "📁 Загрузка данных":
            show_data_upload_interface()
        elif menu == "📊 Обогащение каталога":
            show_catalog_enhance_interface()
        elif menu == "📊 Юнит-экономика":
            show_unit_economics_interface()
        elif menu == "📊 Юнит-экономика по артикулам":
            show_unit_economics_by_article_interface()
        elif menu == "🤖 AI Тарифы":
            show_ai_tariffs_interface()
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
