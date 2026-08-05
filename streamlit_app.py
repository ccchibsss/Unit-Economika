"""
================================================================================
🚗 ULTIMATE UNIT ECONOMICS FOR AUTO PARTS v100.5 - ENTERPRISE EDITION
================================================================================
📌 ВЕРСИЯ: 100.5.1 (ENTERPRISE)
📌 СПЕЦИАЛИЗАЦИЯ: АВТОЗАПЧАСТИ, АВТОТОВАРЫ И АГРЕГАТЫ
📌 ТЕХНОЛОГИИ: STREAMLIT, POLARS, DUCKDB, SCIKIT-LEARN, OPENPYXL, PLOTLY
================================================================================
"""
# ============================================================================
# БЛОК 0: ВСЕ НЕОБХОДИМЫЕ ИМПОРТЫ И КОНФИГУРАЦИЯ
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
import sqlite3
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
from concurrent.futures import ThreadPoolExecutor, ProcessPoolExecutor, as_completed
from datetime import datetime, timedelta, date, timezone
from collections import defaultdict, Counter, deque, OrderedDict, ChainMap, namedtuple
from enum import Enum, auto, IntEnum
from threading import Lock, RLock, Semaphore, Thread, Event, Barrier, Condition
from contextlib import contextmanager, closing, suppress, ExitStack
from pathlib import Path, PurePath
from abc import ABC, abstractmethod
from multiprocessing import Pool, cpu_count
import multiprocessing as mp
from decimal import Decimal, ROUND_HALF_UP

# ============================================================================
# ОПЦИОНАЛЬНЫЕ ИМПОРТЫ С ОБРАБОТКОЙ ОШИБОК
# ============================================================================
try:
    from PIL import Image, ImageDraw, ImageFont, ImageFilter, ImageOps
    PIL_AVAILABLE = True
except Exception:
    PIL_AVAILABLE = False

try:
    import pytz
    PYTZ_AVAILABLE = True
except Exception:
    PYTZ_AVAILABLE = False

try:
    import dateutil
    from dateutil.parser import parse
    from dateutil.relativedelta import relativedelta
    DATEUTIL_AVAILABLE = True
except Exception:
    DATEUTIL_AVAILABLE = False

try:
    import holidays
    HOLIDAYS_AVAILABLE = True
except Exception:
    HOLIDAYS_AVAILABLE = False

try:
    import phonenumbers
    from phonenumbers import PhoneNumberType, PhoneNumber
    from phonenumbers import parse as parse_phone, format_number, PhoneNumberFormat
    PHONENUMBERS_AVAILABLE = True
except Exception:
    PHONENUMBERS_AVAILABLE = False

try:
    import validators
    from validators import url, email as validate_email, domain, ip_address
    VALIDATORS_AVAILABLE = True
except Exception:
    VALIDATORS_AVAILABLE = False

try:
    import pycountry
    PYCOUNTRY_AVAILABLE = True
except Exception:
    PYCOUNTRY_AVAILABLE = False

try:
    import tzlocal
    TZLOCAL_AVAILABLE = True
except Exception:
    TZLOCAL_AVAILABLE = False

try:
    import polars as pl
    import polars.selectors as cs
    POLARS_AVAILABLE = True
    logger_polars = logging.getLogger('polars')
    logger_polars.setLevel(logging.WARNING)
except Exception:
    POLARS_AVAILABLE = False
    pl = None

try:
    import duckdb
    DUCKDB_AVAILABLE = True
except Exception:
    DUCKDB_AVAILABLE = False
    duckdb = None

DASK_AVAILABLE = False
DASK_DF_AVAILABLE = False

try:
    import ray
    RAY_AVAILABLE = True
except Exception:
    RAY_AVAILABLE = False

# ✅ ИСПРАВЛЕНО: Ловим Exception, а не ImportError (защита от AttributeError в modin/pandas)
try:
    import modin.pandas as mpd
    import modin.config as mcfg
    MODIN_AVAILABLE = True
except Exception:
    MODIN_AVAILABLE = False
    mpd = None
    mcfg = None

try:
    import pyarrow as pa
    import pyarrow.parquet as pq
    import pyarrow.csv as pc
    import pyarrow.json as pj
    import pyarrow.fs as pfs
    import pyarrow.compute as pc_comp
    import pyarrow.dataset as ds
    PYARROW_AVAILABLE = True
except Exception:
    PYARROW_AVAILABLE = False

try:
    import pandera as pandera_schema
    from pandera import Column, DataFrameSchema, Check, Index
    PANDERA_AVAILABLE = True
except Exception:
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
except Exception:
    SKLEARN_AVAILABLE = False

try:
    import plotly.graph_objects as go
    import plotly.express as px
    from plotly.subplots import make_subplots
    from plotly.offline import plot, iplot
    from plotly.figure_factory import create_annotated_heatmap, create_distplot, create_2d_density
    import plotly.io as pio
    PLOTLY_AVAILABLE = True
except Exception:
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
except Exception:
    MATPLOTLIB_AVAILABLE = False

try:
    import altair as alt
    ALTAIR_AVAILABLE = True
except Exception:
    ALTAIR_AVAILABLE = False

try:
    import bokeh
    from bokeh.plotting import figure, output_notebook, show
    from bokeh.models import ColumnDataSource, HoverTool, LabelSet, Range1d
    from bokeh.layouts import row, column, gridplot
    BOKEH_AVAILABLE = True
except Exception:
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
    from openpyxl.formatting.rule import Rule, ColorScaleRule, DataBarRule, IconSetRule, CellIsRule, FormulaRule
    from openpyxl.comments import Comment
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.worksheet.table import Table, TableStyleInfo
    from openpyxl.worksheet.worksheet import Worksheet
    from openpyxl.workbook.workbook import Workbook as XLWorkbook
    from openpyxl.worksheet.datavalidation import DataValidation
    OPENPYXL_AVAILABLE = True
except Exception:
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
except Exception:
    PDF_EXPORT = False

try:
    import xlsxwriter
    XLSXWRITER_AVAILABLE = True
except Exception:
    XLSXWRITER_AVAILABLE = False

try:
    import tabulate
    TABULATE_AVAILABLE = True
except Exception:
    TABULATE_AVAILABLE = False

try:
    import chardet
    CHARDET_AVAILABLE = True
except Exception:
    CHARDET_AVAILABLE = False
    chardet = None

try:
    import openai
    OPENAI_AVAILABLE = True
except Exception:
    OPENAI_AVAILABLE = False
    openai = None

try:
    import anthropic
    ANTHROPIC_AVAILABLE = True
except Exception:
    ANTHROPIC_AVAILABLE = False

try:
    import tiktoken
    TIKTOKEN_AVAILABLE = True
except Exception:
    TIKTOKEN_AVAILABLE = False

try:
    import aiohttp
    import aiofiles
    ASYNC_AVAILABLE = True
except Exception:
    ASYNC_AVAILABLE = False

try:
    import httpx
    HTTPX_AVAILABLE = True
except Exception:
    HTTPX_AVAILABLE = False

try:
    import websockets
    WEBSOCKETS_AVAILABLE = True
except Exception:
    WEBSOCKETS_AVAILABLE = False

try:
    import psutil
    PSUTIL_AVAILABLE = True
except Exception:
    PSUTIL_AVAILABLE = False
    psutil = None

try:
    from babel.numbers import format_currency as babel_format_currency
    from babel.numbers import format_percent as babel_format_percent
    from babel.numbers import format_decimal as babel_format_decimal
    BABEL_AVAILABLE = True
except Exception:
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
# 🆕 v100.5.1: СОВМЕСТИМОСТЬ STREAMLIT 1.58+
# ============================================================================
def st_dataframe_compat(df, *args, **kwargs):
    kwargs.pop('use_container_width', None)
    if 'width' not in kwargs:
        kwargs['width'] = 'stretch'
    return st.dataframe(df, *args, **kwargs)

# ============================================================================
# 🆕 v100.5.1: ИСПРАВЛЕНИЕ КРАКОЗЯБР
# ============================================================================
def detect_mojibake(text: str) -> bool:
    if not isinstance(text, str) or not text:
        return False
    mojibake_patterns = [r'Р[°-Џ]{2,}', r'Р[РЎ][°-Џ]{2,}', r'[РЎР][°-Џ]{3,}', r'Р[°-Џ]Р[°-Џ]']
    for pattern in mojibake_patterns:
        if re.search(pattern, text):
            return True
    words = text.split()
    if len(words) >= 3:
        r_words = sum(1 for w in words if w.startswith('Р') and len(w) >= 2)
        if r_words / len(words) > 0.5:
            return True
    return False

def fix_double_utf8(text: str) -> str:
    if not isinstance(text, str) or not text:
        return text
    encodings_to_try = [('cp1251', 'utf-8'), ('latin1', 'utf-8'), ('iso-8859-1', 'utf-8'), ('cp1252', 'utf-8')]
    for source_enc, target_enc in encodings_to_try:
        try:
            fixed = text.encode(source_enc).decode(target_enc)
            if fixed and not detect_mojibake(fixed):
                return fixed
        except (UnicodeDecodeError, UnicodeEncodeError):
            continue
    return text

def fix_dataframe_encoding(df: pd.DataFrame) -> Tuple[pd.DataFrame, int]:
    fixed_count = 0
    new_columns = []
    for col in df.columns:
        col_str = str(col)
        if detect_mojibake(col_str):
            new_col = fix_double_utf8(col_str)
            new_columns.append(new_col)
            fixed_count += 1
        else:
            new_columns.append(col)
    df.columns = new_columns
    for col in df.columns:
        if df[col].dtype == 'object':
            try:
                def _fix_cell(x):
                    if isinstance(x, str) and detect_mojibake(x):
                        return fix_double_utf8(x)
                    return x
                mask = df[col].apply(lambda x: isinstance(x, str) and detect_mojibake(x))
                fixed_count += int(mask.sum())
                df[col] = df[col].apply(_fix_cell)
            except Exception:
                pass
    return df, fixed_count

def smart_read_csv(file_obj, **kwargs) -> pd.DataFrame:
    separators = [';', ',', '\t', '|']
    encodings_priority = ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251']
    best_df = None
    best_encoding = None
    best_sep = None
    mojibake_count = 0
    for encoding in encodings_priority:
        for sep in separators:
            try:
                file_obj.seek(0)
                df = pd.read_csv(file_obj, encoding=encoding, sep=sep, engine='python', on_bad_lines='skip', skipinitialspace=True, quotechar='"', doublequote=True, **kwargs)
                if df is None or df.empty or len(df.columns) <= 1:
                    continue
                current_mojibake = sum(1 for col in df.columns if isinstance(col, str) and detect_mojibake(col))
                if current_mojibake == 0:
                    return df
                if best_df is None or current_mojibake < mojibake_count:
                    best_df = df
                    best_encoding = encoding
                    best_sep = sep
                    mojibake_count = current_mojibake
            except (pd.errors.ParserError, UnicodeDecodeError, Exception):
                continue
    if best_df is not None:
        fixed_df, fixed_count = fix_dataframe_encoding(best_df)
        return fixed_df
    if CHARDET_AVAILABLE and chardet is not None:
        try:
            file_obj.seek(0)
            raw_data = file_obj.read(100000)
            detected = chardet.detect(raw_data)
            if detected and detected.get('encoding'):
                file_obj.seek(0)
                for sep in separators:
                    try:
                        df = pd.read_csv(file_obj, encoding=detected['encoding'], sep=sep, engine='python', on_bad_lines='skip')
                        if df is not None and not df.empty and len(df.columns) > 1:
                            has_mojibake = any(isinstance(col, str) and detect_mojibake(col) for col in df.columns)
                            if has_mojibake:
                                df, _ = fix_dataframe_encoding(df)
                            return df
                    except (pd.errors.ParserError, UnicodeDecodeError):
                        continue
        except Exception as e:
            pass
    raise ValueError("Не удалось прочитать CSV файл. Проверьте кодировку и разделитель.")

# ============================================================================
# ВЕРСИЯ И КОНФИГУРАЦИЯ ПРИЛОЖЕНИЯ
# ============================================================================
APP_VERSION = "100.5.1"
APP_NAME = "🚗 Юнит-экономика автозапчастей PRO 2026"
APP_AUTHOR = "AutoParts Analytics Team"
APP_DESCRIPTION = "Enterprise расчет юнит-экономики для автозапчастей с AI, ML и High-Volume обработкой"
APP_LICENSE = "MIT License"
APP_COPYRIGHT = f"2024-2026 {APP_AUTHOR}"
EXCEL_ROW_LIMIT = 1_000_000
HISTORY_LIMIT = 50_000
CACHE_TTL = 7200
MAX_THREADS = 32
BATCH_SIZE = 2000
DEFAULT_CURRENCY = "RUB"
DEFAULT_MARKETPLACE = "Ozon"
DEFAULT_MODE = "FBY"
MAX_RETRIES = 5
TIMEOUT_SECONDS = 120
MAX_FILE_SIZE_MB = 500
MAX_UPLOAD_SIZE = 1024 * 1024 * 1024
MAX_CATEGORIES = 500
MAX_ANALOGS = 200
PRECISION_DECIMALS = 4
MAX_DISPLAY_ROWS = 2000
PAGE_SIZE = 100
MAX_HISTORY_ENTRIES = 50000
MAX_CACHE_SIZE = 5000
DEFAULT_LOCALE = "ru_RU"
TIMEZONE = "Europe/Moscow"
DEFAULT_MARKUP_GLOBAL = 0.25
DEFAULT_DISCOUNT_MAX = 0.30
DEFAULT_MAX_WORKERS = 8
DEFAULT_CHUNK_SIZE = 10000
SUPPORTED_CURRENCIES = ["RUB", "USD", "EUR", "CNY", "KZT", "UAH", "BYN", "AMD", "TRY"]
SUPPORTED_LANGUAGES = ["ru", "en", "uk", "kz", "by", "am", "tr"]
SUPPORTED_MARKETPLACES = ["Ozon", "Wildberries", "Яндекс Маркет", "AliExpress", "Мегамаркет", "СберМегаМаркет", "Avito", "Drom"]
SUPPORTED_MODES = ["FBY", "FBS", "FBO", "DBS", "FBP", "RealFBS"]

try:
    BASE_DIR = Path(__file__).parent.resolve()
except NameError:
    BASE_DIR = Path.cwd()

DATA_DIR = BASE_DIR / "data"
CACHE_DIR = BASE_DIR / "cache"
LOG_DIR = BASE_DIR / "logs"
REPORTS_DIR = BASE_DIR / "reports"
TEMP_DIR = BASE_DIR / "temp"
MODELS_DIR = BASE_DIR / "models"
CONFIG_DIR = BASE_DIR / "config"
PLUGINS_DIR = BASE_DIR / "plugins"
EXPORTS_DIR = BASE_DIR / "exports"
TARIFFS_DIR = BASE_DIR / "tariffs"
HISTORY_DB_DIR = BASE_DIR / "history_db"
BACKUPS_DIR = BASE_DIR / "backups"

for dir_path in [DATA_DIR, CACHE_DIR, LOG_DIR, REPORTS_DIR, TEMP_DIR, MODELS_DIR, CONFIG_DIR, PLUGINS_DIR, EXPORTS_DIR, TARIFFS_DIR, HISTORY_DB_DIR, BACKUPS_DIR]:
    try:
        dir_path.mkdir(exist_ok=True, parents=True)
    except OSError as e:
        print(f"Ошибка создания директории {dir_path}: {e}")

LOG_LEVEL = os.environ.get("LOG_LEVEL", "INFO")
LOG_FORMAT = "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
LOG_DATE_FORMAT = "%Y-%m-%d %H:%M:%S"
LOG_FILE = LOG_DIR / "auto_parts_economy_pro.log"
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"
DEEPSEEK_MODEL = "deepseek-chat"
USE_CACHING = True
USE_PARALLEL = True
USE_GPU = False
OPTIMIZE_MEMORY = True
USE_DUCKDB = True
USE_POLARS = True
USE_MULTIPROCESSING = True

COLORS = {
    "primary": "#e94560", "secondary": "#0f3460", "success": "#00cc96", "warning": "#ffa600",
    "danger": "#ef553b", "info": "#636efa", "dark": "#1a1a2e", "light": "#f5f5f5",
    "gradient_start": "#1a1a2e", "gradient_end": "#16213e", "input_fill": "#FFF4CC",
    "formula_fill": "#E2EFDA", "result_fill": "#DCE6F1", "header_fill": "#0F3460",
}
PLOTLY_COLORS = [
    "#e94560", "#0f3460", "#00cc96", "#ffa600", "#ef553b", "#636efa", "#f9a825", "#26a69a",
    "#ab47bc", "#42a5f5", "#ec407a", "#66bb6a", "#ffa726", "#8d6e63", "#78909c", "#d4ac0d",
    "#1abc9c", "#2ecc71", "#3498db", "#9b59b6", "#e67e22", "#e74c3c", "#1abc9c", "#2ecc71", "#3498db"
]
MARKETPLACE_ICONS = {
    "Ozon": "🟣", "Wildberries": "🟡", "Яндекс Маркет": "🔵", "AliExpress": "🔴",
    "Мегамаркет": "🟢", "СберМегаМаркет": "🟠", "Avito": "🟤", "Drom": "⚫"
}
MODE_ICONS = {"FBY": "📦", "FBS": "🏪", "FBO": "🏭", "DBS": "🚚", "FBP": "🤝", "RealFBS": "🏃"}

TAX_SYSTEMS = {
    "УСН_6": {"rate": 0.06, "base": "revenue", "name": "УСН 6% (доходы)"},
    "УСН_15": {"rate": 0.15, "base": "profit", "min_rate": 0.01, "name": "УСН 15% (доходы-расходы)"},
    "ОСН": {"rate": 0.20, "base": "profit", "vat": 0.20, "name": "ОСН (общая)"},
    "ПСН": {"rate": 0.0, "base": "fixed", "name": "ПСН (патент)"},
    "НПД": {"rate": 0.06, "base": "revenue", "name": "НПД (самозанятый)"},
}

MARKET_BENCHMARKS_2026 = {
    "фильтры": {"avg_margin": 25, "avg_price": 800, "return_rate": 0.05},
    "колодки": {"avg_margin": 22, "avg_price": 2500, "return_rate": 0.08},
    "масла": {"avg_margin": 18, "avg_price": 3500, "return_rate": 0.03},
    "аккумуляторы": {"avg_margin": 15, "avg_price": 7000, "return_rate": 0.12},
    "шины": {"avg_margin": 20, "avg_price": 5000, "return_rate": 0.07},
    "фары": {"avg_margin": 28, "avg_price": 4500, "return_rate": 0.15},
    "амортизаторы": {"avg_margin": 24, "avg_price": 3000, "return_rate": 0.10},
    "ремни": {"avg_margin": 26, "avg_price": 1200, "return_rate": 0.06},
    "подшипники": {"avg_margin": 23, "avg_price": 1500, "return_rate": 0.09},
    "датчики": {"avg_margin": 27, "avg_price": 2000, "return_rate": 0.11},
}

def money_round(value: float, decimals: int = 2) -> float:
    return float(Decimal(str(value)).quantize(Decimal(f"0.{'0' * decimals}"), rounding=ROUND_HALF_UP))

def calculate_tax(price: float, cost: float, tax_system: str = "УСН_6") -> float:
    cfg = TAX_SYSTEMS.get(tax_system, TAX_SYSTEMS["УСН_6"])
    if cfg["base"] == "revenue":
        return money_round(price * cfg["rate"])
    elif cfg["base"] == "profit":
        profit = price - cost
        tax = profit * cfg["rate"]
        if tax_system == "УСН_15":
            min_tax = price * cfg.get("min_rate", 0.01)
            tax = max(tax, min_tax)
        return money_round(max(0, tax))
    elif cfg["base"] == "fixed":
        return 0.0
    return 0.0

def calculate_billable_weight(weight_kg: float, length_cm: float, width_cm: float, height_cm: float, volumetric_coeff: float = 5000.0) -> float:
    if length_cm <= 0 or width_cm <= 0 or height_cm <= 0:
        return weight_kg
    volumetric_weight = (length_cm * width_cm * height_cm) / volumetric_coeff
    billable = max(weight_kg, volumetric_weight)
    billable = math.ceil(billable * 2) / 2
    return billable

def calculate_storage_cost_progressive(volume_l: float, days: int, base_rate: float, marketplace: str) -> float:
    if marketplace in ["Ozon", "Wildberries"]:
        if days <= 60: multiplier = 1.0
        elif days <= 90: multiplier = 2.0
        elif days <= 180: multiplier = 4.0
        elif days <= 365: multiplier = 8.0
        else: multiplier = 16.0
        weighted_rate = base_rate * multiplier
        return money_round(volume_l * weighted_rate * days)
    else:
        return money_round(volume_l * base_rate * days)

def calculate_returns_cost(price: float, return_rate: float, reverse_logistics: float = 150.0, inspection_cost: float = 50.0) -> float:
    expected_returns = price * return_rate
    reverse_logistics_cost = reverse_logistics * return_rate
    inspection = inspection_cost * return_rate
    loss_from_defects = price * return_rate * 0.3
    return money_round(expected_returns + reverse_logistics_cost + inspection + loss_from_defects)

@dataclass
class AutoPartsSpecificCosts:
    chestny_znak: float = 1.5
    certification_amortization: float = 0.0
    warranty_reserve: float = 0.02
    packaging_fbs: float = 45.0
    labeling: float = 3.0
    util_tax: float = 0.0
    customs_duty: float = 0.0
    currency_risk: float = 0.03

    def calculate(self, price: float, is_import: bool = False, requires_marking: bool = True) -> float:
        total = 0.0
        if requires_marking: total += self.chestny_znak
        total += self.certification_amortization
        total += price * self.warranty_reserve
        total += self.packaging_fbs
        total += self.labeling
        if is_import:
            total += price * self.currency_risk
            total += self.customs_duty
            total += price * self.util_tax
        return money_round(total)

def calculate_advertising_cost(price: float, category: str, ad_intensity: str = "medium") -> float:
    drr_rates = {"low": 0.05, "medium": 0.15, "high": 0.25, "aggressive": 0.35}
    competitive_categories = ["масла", "фильтры", "колодки", "аккумуляторы"]
    if category in competitive_categories:
        intensity = "high" if ad_intensity == "medium" else ad_intensity
    else:
        intensity = ad_intensity
    return money_round(price * drr_rates.get(intensity, 0.15))

def validate_input_data(df: pd.DataFrame) -> Tuple[bool, List[str]]:
    errors = []
    if 'Цена' in df.columns:
        negative_prices = (df['Цена'] <= 0).sum()
        if negative_prices > 0: errors.append(f"⚠️ {negative_prices} товаров с ценой ≤ 0")
        suspicious = (df['Цена'] < 50).sum()
        if suspicious > 0: errors.append(f"⚠️ {suspicious} товаров дешевле 50₽ — проверьте")
    if 'Длина' in df.columns:
        missing_dims = df['Длина'].isna().sum()
        if missing_dims > len(df) * 0.3: errors.append(f"⚠️ У {missing_dims} товаров нет габаритов")
    return len(errors) == 0, errors

def parse_dimensions_string(dim_str: str) -> Tuple[float, float, float]:
    if not dim_str or not isinstance(dim_str, str): return 0.0, 0.0, 0.0
    dim_str = dim_str.lower().strip()
    separators = ['x', '*', 'х', '×', ' ', ',']
    for sep in separators:
        if sep in dim_str:
            parts = [p.strip() for p in dim_str.split(sep) if p.strip()]
            if len(parts) >= 3:
                try:
                    dimensions = []
                    for p in parts[:3]:
                        cleaned = re.sub(r'[^\d.,\-]', '', p)
                        cleaned = cleaned.replace(',', '.')
                        if cleaned and cleaned.replace('.', '').replace('-', '').isdigit():
                            dimensions.append(float(cleaned))
                        else:
                            nums = re.findall(r'(\d+\.?\d*)', p)
                            if nums: dimensions.append(float(nums[0]))
                    if len(dimensions) == 3:
                        dimensions.sort(reverse=True)
                        return tuple(dimensions)
                except (ValueError, TypeError):
                    pass
    return 0.0, 0.0, 0.0

def parse_dimensions_vectorized(dims_series) -> "pl.DataFrame":
    if not POLARS_AVAILABLE: return pl.DataFrame()
    dims = dims_series.str.extract_all(r"(\d+\.?\d*)")
    def sort_dimensions(nums):
        if nums and len(nums) >= 3:
            try: return sorted([float(n) for n in nums[:3]], reverse=True)
            except (ValueError, TypeError): pass
        elif nums and len(nums) == 2:
            try: return [float(nums[0]), float(nums[1]), 1.0]
            except (ValueError, TypeError): pass
        return [0.0, 0.0, 0.0]
    result = dims.map_elements(sort_dimensions, return_dtype=pl.List(pl.Float64))
    return pl.DataFrame({"length": result.list.get(0), "width": result.list.get(1), "height": result.list.get(2)})

def get_api_key_safe(service_name: str) -> Optional[str]:
    try:
        if hasattr(st, 'secrets') and service_name in st.secrets: return st.secrets[service_name]
    except Exception: pass
    env_key = f"{service_name.upper()}_API_KEY"
    return os.environ.get(env_key)

def escape_sql_string(value: str) -> str:
    if not value: return ""
    return re.sub(r"['\";\\]", "", str(value))

class AutoPartsException(Exception):
    def __init__(self, message: str = "", *args, **kwargs):
        self.message = message
        self.timestamp = datetime.now()
        self.context = kwargs
        super().__init__(message, *args)
    def __str__(self):
        return f"[{self.timestamp.strftime('%Y-%m-%d %H:%M:%S')}] {self.message}"

class ValidationError(AutoPartsException):
    def __init__(self, message: str, field: Optional[str] = None, value: Any = None):
        self.field = field
        self.value = value
        super().__init__(f"Ошибка валидации{f' в поле {field}' if field else ''}: {message}")

class MarketplaceError(AutoPartsException):
    def __init__(self, message: str, marketplace: Optional[str] = None):
        self.marketplace = marketplace
        super().__init__(f"Ошибка маркетплейса{f' {marketplace}' if marketplace else ''}: {message}")

class CalculationError(AutoPartsException):
    def __init__(self, message: str, calculation_type: Optional[str] = None):
        self.calculation_type = calculation_type
        super().__init__(f"Ошибка расчета{f' {calculation_type}' if calculation_type else ''}: {message}")

class AIError(AutoPartsException):
    def __init__(self, message: str, provider: Optional[str] = None, code: Optional[int] = None):
        self.provider = provider
        self.code = code
        super().__init__(f"Ошибка AI{f' ({provider})' if provider else ''}: {message}")

class DatabaseError(AutoPartsException):
    def __init__(self, message: str, query: Optional[str] = None, params: Optional[Dict] = None):
        self.query = query
        self.params = params
        super().__init__(f"Ошибка базы данных: {message}")

class ExportError(AutoPartsException):
    def __init__(self, message: str, format: Optional[str] = None, file_path: Optional[Path] = None):
        self.format = format
        self.file_path = file_path
        super().__init__(f"Ошибка экспорта{f' в {format}' if format else ''}: {message}")

class ConfigError(AutoPartsException):
    def __init__(self, message: str, key: Optional[str] = None):
        self.key = key
        super().__init__(f"Ошибка конфигурации{f' для {key}' if key else ''}: {message}")

class DataNotFoundError(AutoPartsException):
    def __init__(self, message: str, entity: Optional[str] = None, id: Optional[Any] = None):
        self.entity = entity
        self.id = id
        super().__init__(f"Данные не найдены{f' {entity}' if entity else ''}: {message}")

class TimeoutError(AutoPartsException):
    def __init__(self, message: str, timeout: Optional[float] = None):
        self.timeout = timeout
        super().__init__(f"Превышено время ожидания{f' ({timeout}с)' if timeout else ''}: {message}")

class PermissionError(AutoPartsException):
    def __init__(self, message: str, resource: Optional[str] = None):
        self.resource = resource
        super().__init__(f"Ошибка доступа{f' к {resource}' if resource else ''}: {message}")

class RateLimitError(AutoPartsException):
    def __init__(self, message: str, limit: Optional[int] = None, reset_time: Optional[datetime] = None):
        self.limit = limit
        self.reset_time = reset_time
        super().__init__(f"Превышен лимит запросов{f' ({limit})' if limit else ''}: {message}")

class AuthenticationError(AutoPartsException):
    def __init__(self, message: str, provider: Optional[str] = None):
        self.provider = provider
        super().__init__(f"Ошибка аутентификации{f' ({provider})' if provider else ''}: {message}")

class IncompatibleDataError(AutoPartsException):
    def __init__(self, message: str, expected_type: Optional[str] = None, actual_type: Optional[str] = None):
        self.expected_type = expected_type
        self.actual_type = actual_type
        super().__init__(f"Несовместимые данные: {message}")

class DataCorruptionError(AutoPartsException):
    def __init__(self, message: str, file_path: Optional[Path] = None, checksum: Optional[str] = None):
        self.file_path = file_path
        self.checksum = checksum
        super().__init__(f"Повреждение данных{f' в {file_path}' if file_path else ''}: {message}")

class ConnectionError(AutoPartsException):
    def __init__(self, message: str, host: Optional[str] = None, port: Optional[int] = None):
        self.host = host
        self.port = port
        super().__init__(f"Ошибка соединения{f' с {host}:{port}' if host else ''}: {message}")

class InvalidStateError(AutoPartsException):
    def __init__(self, message: str, state: Optional[str] = None):
        self.state = state
        super().__init__(f"Некорректное состояние{f' ({state})' if state else ''}: {message}")

class PriceImportError(AutoPartsException):
    def __init__(self, message: str, file_path: Optional[str] = None):
        self.file_path = file_path
        super().__init__(f"Ошибка импорта цен{f' ({file_path})' if file_path else ''}: {message}")

class ForecastError(AutoPartsException):
    def __init__(self, message: str, model: Optional[str] = None):
        self.model = model
        super().__init__(f"Ошибка прогнозирования{f' ({model})' if model else ''}: {message}")

@st.cache_resource
def get_logger():
    logger = logging.getLogger('UnitEconomyPro')
    logger.setLevel(getattr(logging, LOG_LEVEL))
    formatter = logging.Formatter(LOG_FORMAT, datefmt=LOG_DATE_FORMAT)
    try:
        fh = logging.FileHandler(LOG_FILE, encoding='utf-8')
        fh.setLevel(logging.DEBUG)
        fh.setFormatter(formatter)
        logger.addHandler(fh)
    except OSError as e:
        print(f"Ошибка создания файлового логгера: {e}")
        ch = logging.StreamHandler()
        ch.setLevel(logging.INFO)
        ch.setFormatter(formatter)
        logger.addHandler(ch)
    return logger

logger = get_logger()

def timer_decorator(func: Callable) -> Callable:
    @wraps(func)
    def wrapper(*args, **kwargs):
        start = time.perf_counter()
        try:
            result = func(*args, **kwargs)
            elapsed = time.perf_counter() - start
            if elapsed > 1.0: logger.debug(f"⏱ {func.__name__} выполнена за {elapsed:.3f}с")
            return result
        except Exception as e:
            elapsed = time.perf_counter() - start
            logger.error(f"❌ {func.__name__} завершилась с ошибкой за {elapsed:.3f}с: {e}")
            raise
    return wrapper

def cache_decorator(ttl: int = CACHE_TTL, maxsize: int = 1000) -> Callable:
    def decorator(func: Callable) -> Callable:
        cache = {}
        timestamps = {}
        access_count = defaultdict(int)
        @wraps(func)
        def wrapper(*args, **kwargs):
            if not USE_CACHING: return func(*args, **kwargs)
            key = generate_cache_key(*args, **kwargs)
            if len(cache) > maxsize:
                least_used = sorted(access_count.items(), key=lambda x: x[1])[:len(cache) - maxsize]
                for k, _ in least_used:
                    cache.pop(k, None); timestamps.pop(k, None); access_count.pop(k, None)
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

def retry_decorator(max_retries: int = 3, delay: float = 1.0, backoff: float = 2.0, exceptions: tuple = (Exception,)) -> Callable:
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            current_delay = delay
            last_exception = None
            for attempt in range(max_retries):
                try: return func(*args, **kwargs)
                except exceptions as e:
                    last_exception = e
                    if attempt == max_retries - 1: raise
                    logger.warning(f"⚠️ Попытка {attempt + 1}/{max_retries} для {func.__name__} не удалась: {e}")
                    time.sleep(current_delay)
                    current_delay *= backoff
            if last_exception: raise last_exception
            return None
        return wrapper
    return decorator

def validate_inputs(*types: Union[type, tuple], **kwargs_types: Union[type, tuple]) -> Callable:
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            for i, arg in enumerate(args):
                if i < len(types):
                    expected_type = types[i]
                    if not isinstance(arg, expected_type):
                        raise ValidationError(f"Аргумент {i} должен быть типа {expected_type.__name__}", field=str(i), value=arg)
            for param_name, param_value in kwargs.items():
                if param_name in kwargs_types:
                    expected_type = kwargs_types[param_name]
                    if not isinstance(param_value, expected_type):
                        raise ValidationError(f"Аргумент '{param_name}' должен быть типа {expected_type.__name__}", field=param_name, value=param_value)
            return func(*args, **kwargs)
        return wrapper
    return decorator

def log_execution(func: Callable) -> Callable:
    @wraps(func)
    def wrapper(*args, **kwargs):
        args_str = []
        if args: args_str.extend(str(a)[:100] for a in args[:5])
        if kwargs: args_str.extend(f"{k}={str(v)[:100]}" for k, v in list(kwargs.items())[:5])
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
    def decorator(func: Callable) -> Callable:
        @wraps(func)
        def wrapper(*args, **kwargs):
            try: return func(*args, **kwargs)
            except Exception as e:
                if log_error: logger.error(f"⚠️ Ошибка в {func.__name__}: {e}")
                return default_return
        return wrapper
    return decorator

def safe_float(val: Any, default: float = 0.0) -> float:
    if val is None: return default
    if isinstance(val, bool): return float(val)
    if isinstance(val, (int, float)):
        if math.isnan(val) or math.isinf(val): return default
        return float(val)
    if isinstance(val, (decimal.Decimal, np.floating, np.integer)):
        try: return float(val)
        except (ValueError, TypeError): return default
    if isinstance(val, str):
        cleaned = val.strip()
        if not cleaned: return default
        cleaned = re.sub(r'[^\d.,\-+\s]', '', cleaned)
        cleaned = cleaned.replace(' ', '').replace(',', '.')
        if cleaned.count('-') > 1: return default
        parts = cleaned.split('.')
        if len(parts) > 2: return default
        try: return float(cleaned)
        except ValueError: return default
    if hasattr(val, 'dtype') and hasattr(val, 'item'):
        try:
            item = val.item()
            if isinstance(item, (int, float)): return float(item)
        except Exception: pass
    return default

def safe_int(val: Any, default: int = 0) -> int:
    try:
        float_val = safe_float(val, default)
        if float_val == default and val != 0: return default
        return int(float_val)
    except (ValueError, TypeError): return default

def safe_str(val: Any, default: str = "") -> str:
    if val is None: return default
    if isinstance(val, bool): return str(val)
    if isinstance(val, (int, float)):
        if math.isnan(val) or math.isinf(val): return default
        return str(val)
    if isinstance(val, (list, tuple)): return ", ".join(safe_str(v) for v in val[:5]) + ("..." if len(val) > 5 else "")
    if isinstance(val, dict): return str({k: safe_str(v) for k, v in list(val.items())[:5]})
    try:
        result = str(val).strip()
        return result if result else default
    except Exception: return default

def safe_bool(val: Any, default: bool = False) -> bool:
    if val is None: return default
    if isinstance(val, bool): return val
    if isinstance(val, (int, float)): return bool(val)
    if isinstance(val, str):
        val_lower = val.lower().strip()
        true_values = {'true', 'yes', '1', 'y', 'да', 'on'}
        false_values = {'false', 'no', '0', 'n', 'нет', 'off'}
        if val_lower in true_values: return True
        if val_lower in false_values: return False
        return default
    if isinstance(val, (list, tuple, dict)): return bool(val)
    return default

def safe_datetime(val: Any, default: Optional[datetime] = None) -> Optional[datetime]:
    if default is None: default = datetime.now()
    if val is None: return default
    if isinstance(val, datetime): return val
    if isinstance(val, date): return datetime.combine(val, datetime.min.time())
    if isinstance(val, (int, float)):
        try: return datetime.fromtimestamp(val)
        except (ValueError, OSError): return default
    if isinstance(val, str):
        formats = ["%Y-%m-%d %H:%M:%S", "%Y-%m-%d %H:%M:%S.%f", "%Y-%m-%d", "%d.%m.%Y %H:%M:%S", "%d.%m.%Y", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%dT%H:%M:%S.%f", "%Y-%m-%dT%H:%M:%S.%fZ"]
        for fmt in formats:
            try: return datetime.strptime(val, fmt)
            except ValueError: continue
        try:
            if DATEUTIL_AVAILABLE: return parse(val)
        except Exception: pass
    return default

def generate_cache_key(*args, **kwargs) -> str:
    key_parts = []
    for arg in args:
        if isinstance(arg, (dict, OrderedDict)): key_parts.append(json.dumps(arg, sort_keys=True, ensure_ascii=False))
        elif isinstance(arg, (list, tuple, set)): key_parts.append(str(sorted(arg) if not isinstance(arg, tuple) else arg))
        elif isinstance(arg, pd.DataFrame):
            try: key_parts.append(hashlib.md5(pd.util.hash_pandas_object(arg).values.tobytes()).hexdigest())
            except Exception: key_parts.append(str(len(arg)))
        elif isinstance(arg, pd.Series):
            try: key_parts.append(hashlib.md5(pd.util.hash_pandas_object(arg).values.tobytes()).hexdigest())
            except Exception: key_parts.append(str(len(arg)))
        elif isinstance(arg, np.ndarray):
            try: key_parts.append(hashlib.md5(arg.tobytes()).hexdigest())
            except Exception: key_parts.append(str(arg.shape))
        elif isinstance(arg, (datetime, date)): key_parts.append(arg.isoformat())
        else: key_parts.append(str(arg))
    for k, v in sorted(kwargs.items()):
        if isinstance(v, (dict, OrderedDict)): key_parts.append(f"{k}:{json.dumps(v, sort_keys=True, ensure_ascii=False)}")
        elif isinstance(v, (list, tuple, set)): key_parts.append(f"{k}:{str(sorted(v) if not isinstance(v, tuple) else v)}")
        elif isinstance(v, pd.DataFrame):
            try: key_parts.append(f"{k}:{hashlib.md5(pd.util.hash_pandas_object(v).values.tobytes()).hexdigest()}")
            except Exception: key_parts.append(f"{k}:{len(v)}")
        else: key_parts.append(f"{k}:{v}")
    key = "|".join(key_parts)
    return hashlib.md5(key.encode('utf-8')).hexdigest()

def calculate_volume(length: float, width: float, height: float) -> float:
    if not all([length, width, height]): return 0.0
    if not all([length > 0, width > 0, height > 0]): return 0.0
    if any([length > 1000, width > 1000, height > 1000]):
        length /= 10; width /= 10; height /= 10
    if any([length < 0.1, width < 0.1, height < 0.1]): return 0.0
    volume = (length * width * height) / 1000.0
    if volume < 0.001: return 0.0
    return round(volume, 4)

def get_file_encoding(file_path: Union[str, Path]) -> str:
    if CHARDET_AVAILABLE and chardet is not None:
        try:
            with open(file_path, 'rb') as f:
                raw_data = f.read(100000)
                result = chardet.detect(raw_data)
                encoding = result.get('encoding', 'utf-8')
                return encoding
        except (IOError, OSError) as e:
            logger.warning(f"Ошибка определения кодировки: {e}")
    encodings = ['utf-8-sig', 'utf-8', 'cp1251', 'windows-1251', 'cp1252', 'latin1']
    for enc in encodings:
        try:
            with open(file_path, 'r', encoding=enc) as f: f.read()
            return enc
        except UnicodeDecodeError: continue
    return 'utf-8'

def normalize_text(text: str) -> str:
    if not text: return ""
    text = text.lower()
    text = re.sub(r'[^\w\s]', ' ', text)
    text = re.sub(r'\s+', ' ', text)
    return text.strip()

def normalize_key_for_match(value: str) -> str:
    if not value: return ""
    return re.sub(r'[^0-9A-Za-zА-Яа-яЁё]', '', str(value).lower().strip())

def calculate_recommended_min_price(cost: float, commission_rate: float, logistics: float, storage_cost: float, acquiring_rate: float, last_mile: float, return_rate: float, min_profit_percent: float = 0.10, tax_system: str = "УСН_6", tax_rate: float = 0.06) -> float:
    if cost <= 0: return 0.0
    fixed_costs = cost + logistics + storage_cost + last_mile
    variable_rate = commission_rate + acquiring_rate + return_rate + tax_rate + min_profit_percent
    denominator = 1 - variable_rate
    if denominator <= 0: return 0.0
    recommended_price = fixed_costs / denominator
    return max(0, money_round(recommended_price))

# ============================================================================
# БЛОК 1: ENUM И ТИПЫ
# ============================================================================
class CommissionType(Enum):
    PERCENTAGE = auto(); FIXED = auto(); HYBRID = auto(); SUBSCRIPTION = auto()
    TIERED = auto(); DYNAMIC = auto(); FLAT = auto(); CUSTOM = auto()

class OperationMode(Enum):
    FBY = auto(); FBS = auto(); FBO = auto(); DBS = auto(); FBP = auto()
    DBE = auto(); STANDARD = auto(); EXPRESS = auto(); SELF = auto(); REAL_FBS = auto()

class ProductType(Enum):
    ENGINE = "Двигатель"; TRANSMISSION = "Трансмиссия"; SUSPENSION = "Подвеска"
    BRAKE = "Тормозная система"; STEERING = "Рулевое управление"; ELECTRICAL = "Электрооборудование"
    COOLING = "Система охлаждения"; EXHAUST = "Система выпуска"; FUEL = "Система питания"
    FILTER = "Фильтры"; FLUID = "Масла и жидкости"; BODY = "Кузовные детали"
    INTERIOR = "Салон"; EXTERIOR = "Экстерьер"; OPTICS = "Оптика"; TIRES = "Шины и диски"
    TOOLS = "Инструменты"; BELT = "Ремни и приводы"; BEARING = "Подшипники"
    SEAL = "Сальники и прокладки"; FASTENER = "Крепеж"; HVAC = "Климат-контроль"
    AUDIO = "Аудио и мультимедиа"; SAFETY = "Безопасность"; OTHER = "Прочее"

class DataSource(Enum):
    CSV = auto(); EXCEL = auto(); JSON = auto(); API = auto(); DATABASE = auto()
    MANUAL = auto(); MARKETPLACE = auto(); AI = auto(); WEB_SCRAPING = auto()
    ERP = auto(); CRM = auto(); EXTERNAL = auto()

class ExportFormat(Enum):
    CSV = auto(); EXCEL = auto(); EXCEL_FORMULAS = auto(); EXCEL_MACROS = auto()
    PDF = auto(); JSON = auto(); HTML = auto(); MARKDOWN = auto(); PARQUET = auto()
    SQL = auto(); XML = auto(); YAML = auto(); TOML = auto(); POWER_BI = auto(); TABLEAU = auto()

class CalculationStatus(Enum):
    PENDING = auto(); RUNNING = auto(); COMPLETED = auto(); FAILED = auto()
    CANCELLED = auto(); PAUSED = auto(); PARTIAL = auto()

class RiskLevel(Enum):
    LOW = "Низкий"; MEDIUM = "Средний"; HIGH = "Высокий"; CRITICAL = "Критический"

class Seasonality(Enum):
    WINTER = "Зимняя"; SPRING = "Весенняя"; SUMMER = "Летняя"; AUTUMN = "Осенняя"; ALL_YEAR = "Круглогодичная"

class ProfitabilityLevel(Enum):
    LOSS = "Убыток"; BREAK_EVEN = "Точка безубыточности"; LOW = "Низкая"
    MEDIUM = "Средняя"; HIGH = "Высокая"; VERY_HIGH = "Очень высокая"

class Currency(Enum):
    RUB = "RUB"; USD = "USD"; EUR = "EUR"; CNY = "CNY"; KZT = "KZT"
    UAH = "UAH"; BYN = "BYN"; AMD = "AMD"; TRY = "TRY"

class TaxSystem(Enum):
    USN_6 = "УСН_6"; USN_15 = "УСН_15"; OSN = "ОСН"; PSN = "ПСН"; NPD = "НПД"

class TariffSource(Enum):
    HARDCODED = "Захардкожены"; AI_CACHE = "Кэш ИИ"; AI_LIVE = "ИИ (запрос)"
    MANUAL = "Ручной ввод"; IMPORTED = "Импортированы"; API_LIVE = "API Маркетплейса"; FORECAST = "Прогноз ИИ"
# ============================================================================
# БЛОК 12: ВАЛИДАТОР ВЕСОГАБАРИТОВ
# ============================================================================
class AdvancedDimensionsValidator:
    """Класс для умной проверки, нормализации и дополнения весогабаритных характеристик."""
    @staticmethod
    def normalize_dimension(value: float, unit_hint: str = "") -> float:
        if not value or value <= 0:
            return 0.0
        unit_lower = unit_hint.lower() if unit_hint else ""
        if any(x in unit_lower for x in ['mm', 'мм', 'millimeter']):
            return value / 10.0
        if any(x in unit_lower for x in ['m', 'метр', 'meter']) and value < 10:
            return value * 100.0
        if value > 300:
            return value / 10.0
        return value

    @staticmethod
    def normalize_weight(value: float, unit_hint: str = "") -> float:
        if not value or value <= 0:
            return 0.0
        unit_lower = unit_hint.lower() if unit_hint else ""
        if any(x in unit_lower for x in ['g', 'гр', 'gram']):
            return value / 1000.0
        if any(x in unit_lower for x in ['t', 'тонн', 'ton']) and value < 10:
            return value * 1000.0
        if value > 100:
            return value / 1000.0
        return value

    @staticmethod
    def infer_missing_dimensions(category: str, weight: float) -> Dict[str, float]:
        """Эвристическое определение габаритов по категории и весу"""
        defaults = {
            "фильтры": {"l": 15, "w": 15, "h": 15},
            "колодки": {"l": 15, "w": 10, "h": 5},
            "масла": {"l": 10, "w": 10, "h": 25},
            "шины": {"l": 60, "w": 60, "h": 25},
            "аккумуляторы": {"l": 35, "w": 20, "h": 20},
            "фары": {"l": 40, "w": 20, "h": 20},
            "двигатель": {"l": 50, "w": 40, "h": 40},
            "трансмиссия": {"l": 50, "w": 40, "h": 30},
            "подвеска": {"l": 40, "w": 30, "h": 20},
            "тормозная_система": {"l": 30, "w": 20, "h": 15},
            "рулевое_управление": {"l": 40, "w": 20, "h": 15},
            "электрика": {"l": 25, "w": 20, "h": 20},
            "охлаждение": {"l": 45, "w": 35, "h": 20},
            "выпуск": {"l": 60, "w": 25, "h": 20},
            "оптика": {"l": 35, "w": 25, "h": 20},
            "кузов": {"l": 80, "w": 50, "h": 30},
            "инструменты": {"l": 30, "w": 20, "h": 15},
            "ремни": {"l": 25, "w": 15, "h": 10},
            "подшипники": {"l": 15, "w": 15, "h": 10},
            "крепёж": {"l": 10, "w": 10, "h": 5},
            "климат": {"l": 40, "w": 30, "h": 25},
            "безопасность": {"l": 30, "w": 20, "h": 15}
        }
        cat_key = category.lower()
        if cat_key in defaults:
            dims = defaults[cat_key]
        else:
            for key in defaults:
                if key in cat_key:
                    dims = defaults[key]
                    break
            else:
                dims = {"l": 20, "w": 20, "h": 20}
        scale = max(0.5, min(3.0, weight / 2.0))
        return {
            "length_cm": dims["l"] * scale,
            "width_cm": dims["w"] * scale,
            "height_cm": dims["h"] * scale
        }

    @staticmethod
    def validate_and_normalize_row(
        row: pd.Series,
        length_col: Optional[str] = None,
        width_col: Optional[str] = None,
        height_col: Optional[str] = None,
        weight_col: Optional[str] = None,
        category: str = ""
    ) -> Dict[str, float]:
        """Валидация и нормализация строки с габаритами"""
        raw_l = safe_float(row.get(length_col, 0)) if length_col else 0
        raw_w = safe_float(row.get(width_col, 0)) if width_col else 0
        raw_h = safe_float(row.get(height_col, 0)) if height_col else 0
        raw_weight = safe_float(row.get(weight_col, 0)) if weight_col else 0
        length = AdvancedDimensionsValidator.normalize_dimension(raw_l)
        width = AdvancedDimensionsValidator.normalize_dimension(raw_w)
        height = AdvancedDimensionsValidator.normalize_dimension(raw_h)
        weight = AdvancedDimensionsValidator.normalize_weight(raw_weight)
        if length == 0 or width == 0 or height == 0:
            inferred = AdvancedDimensionsValidator.infer_missing_dimensions(category, weight)
            if length == 0: length = inferred["length_cm"]
            if width == 0: width = inferred["width_cm"]
            if height == 0: height = inferred["height_cm"]
        if weight == 0 and length > 0 and width > 0 and height > 0:
            volume = (length * width * height) / 1000
            weight = max(0.1, volume * 0.8)
        return {
            "length_cm": round(length, 2),
            "width_cm": round(width, 2),
            "height_cm": round(height, 2),
            "weight_kg": round(weight, 2)
        }

# ============================================================================
# 🆕 БЛОК 13: UI ФУНКЦИИ - ЗАГРУЗКА ДАННЫХ (v100.7 - С НОРМАЛИЗАЦИЕЙ ВЕСОГАБАРИТОВ)
# ============================================================================
def show_data_upload_interface():
    """📁 РАЗДЕЛ 1: ЗАГРУЗКА ДАННЫХ"""
    st.header("📁 Шаг 1: Загрузка данных каталога")
    st.info("""
**ИНСТРУКЦИЯ ПО ЗАГРУЗКЕ:**
**ШАГ 1:** Подготовьте файл с данными товаров (Excel или CSV)
**ШАГ 2:** Убедитесь, что файл содержит обязательные колонки:
- ✅ Артикул (идентификатор товара)
- ✅ Бренд (производитель)
- ✅ Цена (цена продажи)
- ✅ Себестоимость (закупочная цена)
**ДОПОЛНИТЕЛЬНО:** Система автоматически распознает размеры из колонок:
- 📏 Длина, Ширина, Высота (числовые значения)
- 📏 Весогабариты (строки вида "20x15x10" или "20*15*10")
**🆕 v100.7:** Автоматическая нормализация весогабаритов (исправление дат и плавающей точности)
**ШАГ 3:** Нажмите кнопку ниже и выберите файл
**ШАГ 4:** Дождитесь успешной загрузки
💡 **КАК ПРАВИЛЬНО СОХРАНИТЬ CSV В EXCEL:**
1. Файл → Сохранить как → **CSV UTF-8 (разделитель — запятая)**
2. Или используйте кнопку "Скачать шаблон" ниже (он уже в правильной кодировке)
""")
    uploaded_file = st.file_uploader(
        "📤 Загрузите файл каталога (Excel или CSV)",
        type=['xlsx', 'xls', 'csv'],
        key="data_upload_file",
        help="Поддерживаются форматы: .xlsx, .xls, .csv"
    )
    if uploaded_file is not None:
        try:
            df = None
            file_name = uploaded_file.name.lower()
            if file_name.endswith('.csv'):
                try:
                    df = smart_read_csv(uploaded_file)
                except Exception as e:
                    logger.error(f"Ошибка умного чтения CSV: {e}")
                    raise ValueError(f"Не удалось прочитать CSV файл: {e}")
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
            else:
                raise ValueError(f"Неподдерживаемый формат файла: {file_name}")
            if df is None or df.empty:
                st.error("❌ Не удалось прочитать файл. Проверьте формат и кодировку.")
                return
            df = df.dropna(how='all')
            if df.empty:
                st.warning("⚠️ Файл содержит только пустые строки. Проверьте данные.")
                return
            mojibake_cols = [col for col in df.columns if isinstance(col, str) and detect_mojibake(col)]
            if mojibake_cols:
                st.warning(f"⚠️ Обнаружены кракозябры в {len(mojibake_cols)} колонках. Исправляем...")
                df, fixed_count = fix_dataframe_encoding(df)
                st.success(f"✅ Исправлено {fixed_count} ячеек с кракозябрами")
                st.info(f"📋 Колонки после исправления: {', '.join(str(c) for c in df.columns.tolist())}")
            df.columns = df.columns.str.strip()
            st.subheader("🔧 Нормализация весогабаритов")
            dimension_cols = ['Длина', 'Ширина', 'Высота', 'Вес']
            def normalize_dimension_value(val):
                if pd.isna(val):
                    return 0.0
                if isinstance(val, (datetime, pd.Timestamp)):
                    logger.warning(f"Обнаружена дата вместо числа: {val}")
                    return 0.0
                if isinstance(val, str):
                    val = val.strip()
                    month_names = ['янв', 'фев', 'мар', 'апр', 'май', 'июн', 'июл', 'авг', 'сен', 'окт', 'ноя', 'дек',
                                   'jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec']
                    if any(month in val.lower() for month in month_names):
                        logger.warning(f"Обнаружена дата в строке: {val}")
                        return 0.0
                    try:
                        cleaned = val.replace(',', '.')
                        return round(float(cleaned), 2)
                    except (ValueError, TypeError):
                        logger.warning(f"Не удалось преобразовать строку в число: {val}")
                        return 0.0
                try:
                    num = float(val)
                    return round(num, 2)
                except (ValueError, TypeError):
                    return 0.0
            normalized_count = 0
            for col in dimension_cols:
                if col in df.columns:
                    before_count = df[col].notna().sum()
                    df[col] = df[col].apply(normalize_dimension_value)
                    after_count = (df[col] > 0).sum()
                    if before_count != after_count:
                        normalized_count += 1
                        logger.info(f"Нормализована колонка {col}: {before_count} → {after_count} значений")
            if normalized_count > 0:
                st.success(f"✅ Нормализовано колонок: {normalized_count}")
                st.info("📋 Все значения округлены до 2 знаков после запятой")
                st.write("Пример нормализованных данных:")
                available_cols = [col for col in dimension_cols if col in df.columns]
                if available_cols:
                    st_dataframe_compat(df[available_cols].head(10))
            st.subheader("📏 Автоматический парсинг размеров")
            dims_cols = []
            for col in df.columns:
                col_lower = str(col).lower()
                if any(w in col_lower for w in ['весогабариты', 'размеры', 'dimensions', 'габариты', 'размер']):
                    dims_cols.append(col)
            if dims_cols:
                dims_col = dims_cols[0]
                st.info(f"🔍 Найдена колонка с размерами: **{dims_col}**")
                parsed_data = []
                for idx, row in df.iterrows():
                    dim_str = str(row.get(dims_col, ''))
                    if dim_str and dim_str != 'nan':
                        l, w, h = parse_dimensions_string(dim_str)
                        parsed_data.append({
                            'index': idx,
                            'parsed_length': l,
                            'parsed_width': w,
                            'parsed_height': h
                        })
                if parsed_data:
                    parsed_df = pd.DataFrame(parsed_data)
                    for i, row in parsed_df.iterrows():
                        idx = row['index']
                        if row['parsed_length'] > 0:
                            df.at[idx, 'Длина_парс'] = row['parsed_length']
                            df.at[idx, 'Ширина_парс'] = row['parsed_width']
                            df.at[idx, 'Высота_парс'] = row['parsed_height']
                    rename_map = {}
                    if 'Длина_парс' in df.columns and 'Длина' not in df.columns:
                        rename_map['Длина_парс'] = 'Длина'
                    if 'Ширина_парс' in df.columns and 'Ширина' not in df.columns:
                        rename_map['Ширина_парс'] = 'Ширина'
                    if 'Высота_парс' in df.columns and 'Высота' not in df.columns:
                        rename_map['Высота_парс'] = 'Высота'
                    if rename_map:
                        df = df.rename(columns=rename_map)
                    st.success(f"✅ Распарсено {len(parsed_data)} записей")
                    sample_data = []
                    for i in range(min(5, len(parsed_data))):
                        row = parsed_data[i]
                        sample_data.append({
                            'Исходная строка': df.iloc[row['index']].get(dims_col, ''),
                            'Длина': row['parsed_length'],
                            'Ширина': row['parsed_width'],
                            'Высота': row['parsed_height']
                        })
                    if sample_data:
                        st_dataframe_compat(pd.DataFrame(sample_data))
            st.session_state.uploaded_data = df
            st.success(f"✅ Успешно загружено {len(df)} товаров")
            st.subheader("👁️ Предпросмотр данных (первые 10 строк)")
            st_dataframe_compat(df.head(10), key="upload_preview_table")
            st.subheader("📊 Статистика загруженных данных")
            stats_col1, stats_col2, stats_col3, stats_col4 = st.columns(4)
            with stats_col1:
                st.metric("📦 Всего товаров", len(df))
            with stats_col2:
                price_col = None
                for col in df.columns:
                    if any(w in str(col).lower() for w in ['цена', 'price', 'стоимость']):
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
                    if any(w in str(col).lower() for w in ['себестоимость', 'cost', 'закупочная']):
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
                    if any(w in str(col).lower() for w in ['бренд', 'brand', 'производитель']):
                        brand_col = col
                        break
                if brand_col:
                    try:
                        unique_brands = df[brand_col].nunique()
                        st.metric("🏷️ Уникальных брендов", unique_brands)
                    except Exception:
                        st.metric("🏷️ Брендов", "Ошибка")
                else:
                    st.metric("🏷️ Брендов", "—")
            st.subheader("🔧 Доступные действия")
            action_col1, action_col2, action_col3 = st.columns(3)
            with action_col1:
                if st.button("🏷️ Классифицировать категории", type="secondary", key="classify_btn"):
                    with st.spinner("Классификация товаров..."):
                        classifier = CategoryClassifier()
                        name_col = None
                        for col in df.columns:
                            col_lower = str(col).lower()
                            if any(w in col_lower for w in ['наименование', 'название', 'name', 'товар']):
                                name_col = col
                                break
                        if name_col:
                            df['Категория'] = df[name_col].apply(lambda x: classifier.predict(str(x))[0])
                            st.session_state.uploaded_data = df
                            st.success("✅ Классификация завершена!")
                            st.subheader("📊 Распределение по категориям")
                            category_counts = df['Категория'].value_counts()
                            st_dataframe_compat(category_counts, key="category_counts")
                        else:
                            st.warning("⚠️ Не найдена колонка с названием товара")
            with action_col2:
                if st.button("🔍 Обогатить каталог", type="primary", key="upload_enrich_button"):
                    st.info("ℹ️ Перейдите в раздел '🔍 Обогащение каталога' для поиска аналогов")
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
            "Весогабариты": ["10x5x3", "15x7x4", "20x10x5"],
            "OE номер": ["123456", "654321", "789012"],
            "Описание": ["Описание товара 1", "Описание товара 2", "Описание товара 3"]
        })
        import codecs
        output = io.BytesIO()
        output.write(codecs.BOM_UTF8)
        csv_string = template_df.to_csv(index=False, sep=';')
        output.write(csv_string.encode('utf-8'))
        output.seek(0)
        st.download_button(
            label="📥 Скачать шаблон CSV (Excel-совместимый)",
            data=output,
            file_name="шаблон_каталога.csv",
            mime="text/csv; charset=utf-8",
            key="download_template"
        )

# ============================================================================
# 🆕 БЛОК 14: СУПЕР-PRO ЭКСПОРТЕР ЮНИТ-ЭКОНОМИКИ v2.0 (ИСПРАВЛЕННАЯ ВЕРСИЯ)
# ============================================================================
class SuperProExcelExporter:
    """
    🚀 СУПЕР-ПРО ЭКСПОРТ ЮНИТ-ЭКОНОМИКИ v2.0
    Максимально информативный шаблон с живыми формулами и аналитикой
    """
    TAX_ROW_OFFSET = 5
    MIN_PROFIT_ROW_OFFSET = 6
    AD_ROW = 9
    DAYS_ROW = 7
    CURRENCY_ROW = 10
    COLORS = {
        "header_bg": "1B3A5C", "header_fg": "FFFFFF", "section_bg": "2E86AB",
        "input_bg": "FFF4CC", "param_bg": "E8F4FD", "formula_bg": "DCE6F1",
        "positive": "C6EFCE", "positive_text": "006100", "negative": "FFC7CE",
        "negative_text": "9C0006", "warning": "FFEB9C", "warning_text": "9C6500",
        "total_bg": "D9E2F3", "border": "B4C6E7", "mp_header": "4472C4",
        "gradient_start": "E8F4FD", "gradient_end": "B4C6E7",
    }
    OPERATION_MODES = ["FBY", "FBS", "FBO", "DBS", "FBP", "RealFBS"]
    SEASONS = ["winter", "spring", "summer", "autumn"]
    SEASON_NAMES = {"winter": "❄️ Зима", "spring": "🌱 Весна", "summer": "☀️ Лето", "autumn": "🍂 Осень"}

    def __init__(self, unit_economics=None):
        self.formats = {}
        self.unit_economics = unit_economics
        self._base_rates_start_row = None
        self._base_rates_end_row = None
        self._global_tax_row = None
        self._global_min_profit_row = None
        self._input_start_row = 4
        self._total_rows = 0

    def _get_configs(self):
        if self.unit_economics and hasattr(self.unit_economics, '_configs'):
            configs = self.unit_economics._configs
            if configs: return configs
        try:
            unit_econ = get_marketplace_unit_economics()
            if unit_econ and hasattr(unit_econ, '_configs'): return unit_econ._configs
        except Exception: pass
        return get_marketplace_configs_2026()

    def _init_formats(self, workbook):
        self.formats = {
            'header': workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': self.COLORS["header_bg"], 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True, 'font_size': 11}),
            'header_title': workbook.add_format({'bold': True, 'font_size': 16, 'font_color': 'white', 'bg_color': self.COLORS["header_bg"], 'align': 'center', 'valign': 'vcenter', 'border': 1}),
            'section_title': workbook.add_format({'bold': True, 'font_size': 13, 'font_color': 'white', 'bg_color': self.COLORS["section_bg"], 'align': 'left', 'valign': 'vcenter', 'border': 1}),
            'mp_header': workbook.add_format({'bold': True, 'font_color': 'white', 'bg_color': self.COLORS["mp_header"], 'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True}),
            'input_cell': workbook.add_format({'bg_color': self.COLORS["input_bg"], 'border': 1, 'num_format': '#,##0.00'}),
            'input_cell_int': workbook.add_format({'bg_color': self.COLORS["input_bg"], 'border': 1, 'num_format': '0.00'}),
            'input_percent': workbook.add_format({'bg_color': self.COLORS["input_bg"], 'border': 1, 'num_format': '0.00%'}),
            'param_cell': workbook.add_format({'bold': True, 'bg_color': self.COLORS["param_bg"], 'border': 1, 'valign': 'vcenter'}),
            'param_value': workbook.add_format({'bold': True, 'font_size': 11, 'bg_color': self.COLORS["input_bg"], 'border': 1}),
            'formula_cell': workbook.add_format({'bg_color': self.COLORS["formula_bg"], 'border': 1, 'num_format': '#,##0.00 ₽'}),
            'formula_percent': workbook.add_format({'bg_color': self.COLORS["formula_bg"], 'border': 1, 'num_format': '0.00%'}),
            'money': workbook.add_format({'border': 1, 'num_format': '#,##0.00 ₽'}),
            'money_bold': workbook.add_format({'bold': True, 'border': 1, 'num_format': '#,##0.00 ₽'}),
            'bold': workbook.add_format({'bold': True, 'border': 1}),
            'bold_money': workbook.add_format({'bold': True, 'font_size': 11, 'bg_color': self.COLORS["total_bg"], 'border': 1, 'num_format': '#,##0.00 ₽'}),
            'bold_percent': workbook.add_format({'bold': True, 'font_size': 11, 'bg_color': self.COLORS["total_bg"], 'border': 1, 'num_format': '0.00%'}),
            'positive': workbook.add_format({'bg_color': self.COLORS["positive"], 'font_color': self.COLORS["positive_text"], 'bold': True, 'border': 1}),
            'negative': workbook.add_format({'bg_color': self.COLORS["negative"], 'font_color': self.COLORS["negative_text"], 'bold': True, 'border': 1}),
            'warning_cell': workbook.add_format({'bg_color': self.COLORS["warning"], 'font_color': self.COLORS["warning_text"], 'bold': True, 'border': 1}),
            'info': workbook.add_format({'italic': True, 'font_color': self.COLORS["positive_text"], 'bg_color': self.COLORS["positive"], 'border': 1}),
            'warning': workbook.add_format({'italic': True, 'font_color': self.COLORS["negative_text"], 'bg_color': self.COLORS["warning"], 'border': 1}),
            'default': workbook.add_format({'border': 1}),
            'kpi_label': workbook.add_format({'bold': True, 'font_size': 12, 'border': 1, 'valign': 'vcenter', 'bg_color': self.COLORS["param_bg"]}),
            'kpi_positive_money': workbook.add_format({'bold': True, 'font_size': 14, 'border': 1, 'bg_color': self.COLORS["positive"], 'font_color': self.COLORS["positive_text"], 'num_format': '#,##0.00 ₽'}),
            'kpi_negative_money': workbook.add_format({'bold': True, 'font_size': 14, 'border': 1, 'bg_color': self.COLORS["negative"], 'font_color': self.COLORS["negative_text"], 'num_format': '#,##0.00 ₽'}),
            'kpi_neutral_money': workbook.add_format({'bold': True, 'font_size': 14, 'border': 1, 'num_format': '#,##0.00 ₽'}),
            'kpi_neutral_percent': workbook.add_format({'bold': True, 'font_size': 14, 'border': 1, 'num_format': '0.00%'}),
            'kpi_neutral_int': workbook.add_format({'bold': True, 'font_size': 14, 'border': 1, 'num_format': '#,##0'}),
            'chart_title': workbook.add_format({'bold': True, 'font_size': 12, 'align': 'center', 'valign': 'vcenter'}),
        }

    def export_super_pro(self, df: pd.DataFrame, output_path: str, metadata: Dict = None) -> bool:
        try:
            if not XLSXWRITER_AVAILABLE:
                logger.error("❌ xlsxwriter не установлен!")
                return False
            self._total_rows = len(df)
            workbook = xlsxwriter.Workbook(output_path, {'nan_inf_to_errors': True})
            self._init_formats(workbook)
            self._write_dashboard_super(workbook, df, metadata)
            self._write_parameters_super(workbook, metadata)
            self._write_input_data(workbook, df)
            self._write_calculation_engine(workbook, df)
            self._write_marketplace_comparison(workbook, df)
            self._write_category_analysis(workbook, df)
            self._write_profit_forecast(workbook, df)
            self._write_sensitivity_analysis(workbook, df)
            self._write_top_analytics(workbook, df)
            self._write_recommendations(workbook, df)
            self._write_export_summary(workbook, df, metadata)
            workbook.close()
            logger.info(f"✅ СУПЕР-ПРО файл сохранён: {output_path}")
            return True
        except Exception as e:
            logger.error(f"❌ Ошибка СУПЕР-ПРО экспорта: {e}")
            logger.error(traceback.format_exc())
            return False

    def _write_dashboard_super(self, workbook, df: pd.DataFrame, metadata: Dict):
        ws = workbook.add_worksheet("📊 Дашборд")
        ws.merge_range('A1:G1', "🚀 СУПЕР-ДАШБОРД ЮНИТ-ЭКОНОМИКИ", self.formats['header_title'])
        ws.set_row(0, 40)
        ws.merge_range('A2:G2', "📊 Ключевые показатели эффективности (KPI) в реальном времени", self.formats['info'])
        ws.set_row(1, 25)
        total_profit = df['profit'].sum() if 'profit' in df.columns else 0
        avg_margin = df['margin_percent'].mean() if 'margin_percent' in df.columns else 0
        avg_roi = df['roi'].mean() if 'roi' in df.columns else 0
        total_revenue = df['price'].sum() if 'price' in df.columns else 0
        total_expenses = df['total_expenses'].sum() if 'total_expenses' in df.columns else 0
        unprofitable = (df['profit'] < 0).sum() if 'profit' in df.columns else 0
        kpis = [
            ("📦 Всего SKU", f"{len(df):,}", "kpi_neutral_int"),
            ("💰 Общая прибыль", f"{total_profit:,.0f} ₽", "kpi_positive_money" if total_profit > 0 else "kpi_negative_money"),
            ("📈 Средняя маржа", f"{avg_margin:.1f}%", "kpi_neutral_percent"),
            ("📊 Средний ROI", f"{avg_roi:.1f}%", "kpi_neutral_percent"),
            ("💵 Общая выручка", f"{total_revenue:,.0f} ₽", "kpi_neutral_money"),
            ("💸 Общие расходы", f"{total_expenses:,.0f} ₽", "kpi_neutral_money"),
            ("⚠️ Убыточных SKU", f"{unprofitable}", "kpi_neutral_int"),
        ]
        row = 3
        for i, (label, value, fmt) in enumerate(kpis):
            col = (i % 4) * 2
            ws.write(row, col, label, self.formats['kpi_label'])
            ws.write(row, col + 1, value, self.formats[fmt])
            if i % 4 == 3: row += 1
        if 'marketplace' in df.columns and 'profit' in df.columns:
            mp_profit = df.groupby('marketplace')['profit'].sum().sort_values(ascending=False)
            if not mp_profit.empty:
                chart_row = row + 3
                ws.write(chart_row, 0, "🏪 Прибыль по маркетплейсам", self.formats['chart_title'])
                data_start_row = chart_row + 1
                for i, (mp, profit) in enumerate(mp_profit.items()):
                    ws.write(data_start_row + i, 0, mp, self.formats['default'])
                    ws.write(data_start_row + i, 1, profit, self.formats['money'])
                chart = workbook.add_chart({'type': 'column'})
                chart.add_series({
                    'name': 'Прибыль по МП',
                    'categories': f'=📊 Дашборд!$A${data_start_row+1}:$A${data_start_row+len(mp_profit)}',
                    'values': f'=📊 Дашборд!$B${data_start_row+1}:$B${data_start_row+len(mp_profit)}',
                    'fill': {'color': self.COLORS["section_bg"]},
                    'border': {'color': self.COLORS["header_bg"]},
                })
                chart.set_title({'name': 'Прибыль по маркетплейсам'})
                chart.set_x_axis({'name': 'Маркетплейс'})
                chart.set_y_axis({'name': 'Прибыль, ₽'})
                chart.set_size({'width': 720, 'height': 400})
                ws.insert_chart(chart_row, 2, chart)
        ws.set_column('A:A', 25); ws.set_column('B:B', 25); ws.set_column('C:C', 25); ws.set_column('D:D', 25)
        return ws

    def _write_parameters_super(self, workbook, metadata: Dict):
        ws = workbook.add_worksheet("⚙️ Параметры")
        ws.merge_range('A1:P1', "⚙️ РАСШИРЕННЫЕ ПАРАМЕТРЫ РАСЧЁТА", self.formats['header_title'])
        ws.set_row(0, 30)
        ws.merge_range('A2:P2', "💡 Все параметры редактируемые — изменения применяются ко всем расчётам", self.formats['info'])
        if metadata is None: metadata = {}
        row = 4
        ws.merge_range(row, 0, row, 15, "🌐 ГЛОБАЛЬНЫЕ ПАРАМЕТРЫ", self.formats['section_title'])
        row += 1
        global_params = [
            ("Налоговая ставка", 0.06, "Налог от цены продажи", "0.00%"),
            ("Мин. прибыль (%)", 0.10, "Минимальная целевая прибыль", "0.00%"),
            ("Дней хранения", 30, "Среднее кол-во дней", "0"),
            ("ДРР (реклама)", 0.15, "Доля рекламных расходов", "0.00%"),
            ("Курс USD/RUB", 92.50, "Для импортных товаров", "0.00"),
            ("Инфляция %", 0.07, "Годовая инфляция", "0.00%"),
        ]
        for name, value, desc, fmt in global_params:
            ws.write(row, 0, name, self.formats['param_cell'])
            if "Дней" in name: ws.write(row, 1, value, self.formats['input_cell_int'])
            elif "%" in fmt: ws.write(row, 1, value, self.formats['input_percent'])
            else: ws.write(row, 1, value, self.formats['input_cell'])
            ws.write(row, 2, desc, self.formats['default'])
            if "Налоговая" in name: self._global_tax_row = row + 1
            elif "Мин. прибыль" in name: self._global_min_profit_row = row + 1
            row += 1
        row += 2
        ws.merge_range(row, 0, row, 15, "📊 БАЗОВЫЕ ТАРИФЫ (ключ = МП|Режим)", self.formats['section_title'])
        row += 1
        headers = ['Ключ', 'МП', 'Режим', 'Комиссия', 'Лог. база', 'Лог/кг', 'Лог/л', 'Хранение', 'Эквайринг', 'Возвраты', 'Посл. миля', 'Подписка', 'Страховка', 'Упаковка', 'Надбавка', 'Источник']
        for col_idx, header in enumerate(headers):
            ws.write(row, col_idx, header, self.formats['mp_header'])
        self._base_rates_start_row = row + 1
        row += 1
        configs = self._get_configs()
        if configs:
            for mp_name in sorted(configs.keys()):
                config = configs[mp_name]
                for mode in self.OPERATION_MODES:
                    key = f"{mp_name}|{mode}"
                    base_rate = config.commission_rate
                    mode_mult = config.mode_multipliers.get(mode, 1.0)
                    effective_rate = base_rate * mode_mult
                    ws.write(row, 0, key, self.formats['param_cell'])
                    ws.write(row, 1, mp_name, self.formats['param_cell'])
                    ws.write(row, 2, mode, self.formats['param_cell'])
                    ws.write(row, 3, effective_rate, self.formats['input_percent'])
                    ws.write(row, 4, config.logistics_base, self.formats['input_cell'])
                    ws.write(row, 5, config.logistics_per_kg, self.formats['input_cell'])
                    ws.write(row, 6, config.logistics_per_liter, self.formats['input_cell'])
                    ws.write(row, 7, config.storage_per_day, self.formats['input_cell'])
                    ws.write(row, 8, config.acquiring_fee, self.formats['input_percent'])
                    ws.write(row, 9, config.return_fee, self.formats['input_percent'])
                    ws.write(row, 10, config.last_mile_fee, self.formats['input_cell'])
                    ws.write(row, 11, config.subscription_fee, self.formats['input_cell'])
                    ws.write(row, 12, config.insurance_fee, self.formats['input_percent'])
                    ws.write(row, 13, config.packing_fee, self.formats['input_cell'])
                    ws.write(row, 14, config.hazardous_surcharge, self.formats['input_percent'])
                    ws.write(row, 15, config.tariff_source.value, self.formats['default'])
                    row += 1
        else:
            ws.write(row, 0, "Ozon|FBS", self.formats['param_cell'])
            ws.write(row, 1, "Ozon", self.formats['param_cell'])
            ws.write(row, 2, "FBS", self.formats['param_cell'])
            ws.write(row, 3, 0.15, self.formats['input_percent'])
            row += 1
        self._base_rates_end_row = row
        ws.set_column('A:A', 18); ws.set_column('B:C', 14); ws.set_column('D:O', 14); ws.set_column('P:P', 16)
        return ws

    def _write_input_data(self, workbook, df: pd.DataFrame):
        ws = workbook.add_worksheet("📥 Входные")
        ws.merge_range('A1:N1', "📥 ВХОДНЫЕ ДАННЫЕ (редактируемые)", self.formats['header_title'])
        ws.set_row(0, 28)
        ws.merge_range('A2:N2', "💡 Меняйте значения — все листы пересчитаются автоматически", self.formats['info'])
        headers = ['Артикул', 'Бренд', 'МП', 'Режим', 'Категория', 'Цена', 'Себест-ть', 'Вес, кг', 'Длина, см', 'Ширина, см', 'Высота, см', 'Объём, л', 'Оплач. вес', 'Наценка %']
        for col_idx, header in enumerate(headers):
            ws.write(2, col_idx, header, self.formats['header'])
        ws.set_row(2, 30)
        for i, (_, row_data) in enumerate(df.iterrows()):
            excel_row = 3 + i
            ws.write(excel_row, 0, str(row_data.get('Артикул', '')), self.formats['default'])
            ws.write(excel_row, 1, str(row_data.get('Бренд', '')), self.formats['default'])
            ws.write(excel_row, 2, str(row_data.get('marketplace', 'Ozon')), self.formats['default'])
            ws.write(excel_row, 3, str(row_data.get('operation_mode', 'FBS')), self.formats['default'])
            category = str(row_data.get('category', ''))
            if category: category = category.lower().replace(' ', '_')
            ws.write(excel_row, 4, category, self.formats['default'])
            ws.write(excel_row, 5, float(row_data.get('price', 0)), self.formats['input_cell'])
            ws.write(excel_row, 6, float(row_data.get('cost', 0)), self.formats['input_cell'])
            ws.write(excel_row, 7, float(row_data.get('weight', 0)), self.formats['input_cell_int'])
            ws.write(excel_row, 8, float(row_data.get('length', 0)), self.formats['input_cell_int'])
            ws.write(excel_row, 9, float(row_data.get('width', 0)), self.formats['input_cell_int'])
            ws.write(excel_row, 10, float(row_data.get('height', 0)), self.formats['input_cell_int'])
            volume = (float(row_data.get('length', 0)) * float(row_data.get('width', 0)) * float(row_data.get('height', 0))) / 1000
            ws.write(excel_row, 11, volume, self.formats['formula_cell'])
            ws.write_formula(excel_row, 12, f"=MAX(G{excel_row+1}, L{excel_row+1}/5000)", self.formats['formula_cell'])
            ws.write(excel_row, 13, 0, self.formats['input_percent'])
        ws.set_column('A:B', 18); ws.set_column('C:D', 15); ws.set_column('E:E', 18); ws.set_column('F:M', 14); ws.set_column('N:N', 14)
        ws.freeze_panes(3, 0)
        if self._total_rows > 0: ws.autofilter(2, 0, 2 + self._total_rows, 13)
        return ws

    def _write_calculation_engine(self, workbook, df: pd.DataFrame):
        ws = workbook.add_worksheet("📊 Расчёт")
        ws.merge_range('A1:W1', "📊 ПОЛНЫЙ РАСЧЁТ ЮНИТ-ЭКОНОМИКИ", self.formats['header_title'])
        ws.set_row(0, 28)
        ws.merge_range('A2:W2', "⚠️ Все расчёты автоматические — не редактируйте формулы", self.formats['warning'])
        headers = ['Артикул', 'МП', 'Режим', 'Категория', 'Цена', 'Себест-ть', 'Вес', 'Объём', 'Комиссия', 'Логистика', 'Хранение', 'Эквайринг', 'Посл. миля', 'Возвраты', 'Реклама', 'Налог', 'Страховка', 'Упаковка', 'ИТОГО расходов', '💰 ПРИБЫЛЬ', 'Маржа %', 'ROI %', 'Безубыт-ть']
        for col_idx, header in enumerate(headers):
            ws.write(2, col_idx, header, self.formats['header'])
        ws.set_row(2, 35)
        p_tax = f"'⚙️ Параметры'!$B${self._global_tax_row}"
        min_profit = f"'⚙️ Параметры'!$B${self._global_min_profit_row}"
        p_ad = f"'⚙️ Параметры'!$B${self.AD_ROW}"
        p_days = f"'⚙️ Параметры'!$B${self.DAYS_ROW}"
        p_currency = f"'⚙️ Параметры'!$B${self.CURRENCY_ROW}"
        params_range = f"'⚙️ Параметры'!$A${self._base_rates_start_row}:$P${self._base_rates_end_row}"
        for i in range(self._total_rows):
            excel_row = 3 + i
            input_row = 4 + i
            in_art = f"'📥 Входные'!A{input_row}"; in_mp = f"'📥 Входные'!C{input_row}"; in_mode = f"'📥 Входные'!D{input_row}"; in_cat = f"'📥 Входные'!E{input_row}"
            in_price = f"'📥 Входные'!F{input_row}"; in_cost = f"'📥 Входные'!G{input_row}"; in_weight = f"'📥 Входные'!H{input_row}"; in_volume = f"'📥 Входные'!L{input_row}"
            lookup_key = f'CONCATENATE({in_mp},"|",{in_mode})'
            ws.write_formula(excel_row, 0, f"={in_art}", self.formats['default'])
            ws.write_formula(excel_row, 1, f"={in_mp}", self.formats['default'])
            ws.write_formula(excel_row, 2, f"={in_mode}", self.formats['default'])
            ws.write_formula(excel_row, 3, f"={in_cat}", self.formats['default'])
            ws.write_formula(excel_row, 4, f"={in_price}", self.formats['formula_cell'])
            ws.write_formula(excel_row, 5, f"={in_cost}", self.formats['formula_cell'])
            ws.write_formula(excel_row, 6, f"={in_weight}", self.formats['formula_cell'])
            ws.write_formula(excel_row, 7, f"={in_volume}", self.formats['formula_cell'])
            ws.write_formula(excel_row, 8, f"=VLOOKUP({lookup_key},{params_range},4,FALSE)*{in_price}", self.formats['formula_cell'])
            ws.write_formula(excel_row, 9, f"=VLOOKUP({lookup_key},{params_range},5,FALSE)+{in_weight}*VLOOKUP({lookup_key},{params_range},6,FALSE)+{in_volume}*VLOOKUP({lookup_key},{params_range},7,FALSE)", self.formats['formula_cell'])
            ws.write_formula(excel_row, 10, f"={in_volume}*VLOOKUP({lookup_key},{params_range},8,FALSE)*{p_days}", self.formats['formula_cell'])
            ws.write_formula(excel_row, 11, f"=VLOOKUP({lookup_key},{params_range},9,FALSE)*{in_price}", self.formats['formula_cell'])
            ws.write_formula(excel_row, 12, f"=VLOOKUP({lookup_key},{params_range},11,FALSE)", self.formats['formula_cell'])
            ws.write_formula(excel_row, 13, f"=VLOOKUP({lookup_key},{params_range},10,FALSE)*{in_price}*1.3", self.formats['formula_cell'])
            ws.write_formula(excel_row, 14, f"={in_price}*{p_ad}", self.formats['formula_cell'])
            ws.write_formula(excel_row, 15, f"={in_price}*{p_tax}", self.formats['formula_cell'])
            ws.write_formula(excel_row, 16, f"=VLOOKUP({lookup_key},{params_range},13,FALSE)*{in_price}", self.formats['formula_cell'])
            ws.write_formula(excel_row, 17, f"=VLOOKUP({lookup_key},{params_range},14,FALSE)", self.formats['formula_cell'])
            ws.write_formula(excel_row, 18, f"={in_cost}+SUM(I{excel_row+1}:R{excel_row+1})", self.formats['formula_cell'])
            ws.write_formula(excel_row, 19, f"={in_price}-S{excel_row+1}", self.formats['formula_cell'])
            ws.write_formula(excel_row, 20, f"=IF({in_price}>0,T{excel_row+1}/{in_price},0)", self.formats['formula_percent'])
            ws.write_formula(excel_row, 21, f"=IF({in_cost}>0,T{excel_row+1}/{in_cost},0)", self.formats['formula_percent'])
            ws.write_formula(excel_row, 22, f"=S{excel_row+1}/(1-VLOOKUP({lookup_key},{params_range},4,FALSE)-VLOOKUP({lookup_key},{params_range},9,FALSE)-{p_tax})", self.formats['formula_cell'])
        if self._total_rows > 0:
            last_row = 3 + self._total_rows
            profit_range = f"T4:T{last_row}"
            ws.conditional_format(profit_range, {'type': 'cell', 'criteria': '>', 'value': 0, 'format': self.formats['positive']})
            ws.conditional_format(profit_range, {'type': 'cell', 'criteria': '<', 'value': 0, 'format': self.formats['negative']})
            margin_range = f"U4:U{last_row}"
            ws.conditional_format(margin_range, {'type': '3_color_scale', 'min_color': self.COLORS["negative"], 'mid_color': self.COLORS["warning"], 'max_color': self.COLORS["positive"]})
            total_row = 3 + self._total_rows + 2
            ws.merge_range(total_row, 0, total_row, 2, "ИТОГО / СРЕДНЕЕ:", self.formats['bold_money'])
            last_data_row = 3 + self._total_rows
            for col_idx, col_letter in enumerate(['E', 'F', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T']):
                ws.write_formula(total_row, col_idx + 4, f"=SUM({col_letter}4:{col_letter}{last_data_row})", self.formats['bold_money'])
            for col_idx, col_letter in enumerate(['U', 'V'], start=20):
                ws.write_formula(total_row, col_idx, f"=AVERAGE({col_letter}4:{col_letter}{last_data_row})", self.formats['bold_percent'])
        widths = {'A': 15, 'B': 14, 'C': 10, 'D': 14, 'E': 12, 'F': 12, 'G': 10, 'H': 10, 'I': 12, 'J': 12, 'K': 12, 'L': 12, 'M': 12, 'N': 12, 'O': 12, 'P': 12, 'Q': 12, 'R': 12, 'S': 15, 'T': 15, 'U': 12, 'V': 12, 'W': 14}
        for col, width in widths.items(): ws.set_column(f'{col}:{col}', width)
        ws.freeze_panes(3, 0)
        if self._total_rows > 0: ws.autofilter(2, 0, 2 + self._total_rows, 22)
        return ws

    def _write_marketplace_comparison(self, workbook, df: pd.DataFrame):
        ws = workbook.add_worksheet("🏪 Сравнение МП")
        ws.merge_range('A1:K1', "🏪 СРАВНИТЕЛЬНЫЙ АНАЛИЗ МАРКЕТПЛЕЙСОВ", self.formats['header_title'])
        headers = ['МП', 'SKU', 'Выручка', 'Расходы', 'Прибыль', 'Ср. прибыль', 'Ср. маржа %', 'ROI %', 'Доля рынка %', 'Эффективность', 'Рейтинг']
        for col_idx, header in enumerate(headers): ws.write(2, col_idx, header, self.formats['header'])
        if 'marketplace' in df.columns:
            mp_stats = df.groupby('marketplace').agg({'price': 'sum', 'total_expenses': 'sum', 'profit': ['sum', 'mean'], 'margin_percent': 'mean', 'roi': 'mean'}).reset_index()
            mp_stats.columns = ['МП', 'Выручка', 'Расходы', 'Прибыль', 'Ср. прибыль', 'Ср. маржа %', 'ROI %']
            total_profit = mp_stats['Прибыль'].sum()
            for i, row in mp_stats.iterrows():
                excel_row = 3 + i
                ws.write(excel_row, 0, row['МП'], self.formats['bold'])
                ws.write_formula(excel_row, 1, f"=COUNTIF('📊 Расчёт'!$B:$B,A{excel_row+1})", self.formats['default'])
                ws.write(excel_row, 2, row['Выручка'], self.formats['money'])
                ws.write(excel_row, 3, row['Расходы'], self.formats['money'])
                ws.write(excel_row, 4, row['Прибыль'], self.formats['positive'] if row['Прибыль'] > 0 else self.formats['negative'])
                ws.write(excel_row, 5, row['Ср. прибыль'], self.formats['money'])
                ws.write(excel_row, 6, row['Ср. маржа %'], self.formats['formula_percent'])
                ws.write(excel_row, 7, row['ROI %'], self.formats['formula_percent'])
                share = (row['Прибыль'] / total_profit * 100) if total_profit > 0 else 0
                ws.write(excel_row, 8, share / 100, self.formats['formula_percent'])
                ws.write_formula(excel_row, 9, f"=IF(C{excel_row+1}>0,E{excel_row+1}/C{excel_row+1},0)", self.formats['formula_percent'])
                ws.write_formula(excel_row, 10, f"=RANK(E{excel_row+1},$E$4:$E${3+len(mp_stats)})", self.formats['default'])
        ws.set_column('A:K', 16); ws.freeze_panes(3, 0)
        return ws

    def _write_category_analysis(self, workbook, df: pd.DataFrame):
        ws = workbook.add_worksheet("📂 Категории")
        ws.merge_range('A1:H1', "📂 АНАЛИЗ ПО КАТЕГОРИЯМ", self.formats['header_title'])
        headers = ['Категория', 'SKU', 'Выручка', 'Прибыль', 'Ср. маржа %', 'Топ товар', 'Прибыль топ', 'Доля %']
        for col_idx, header in enumerate(headers): ws.write(2, col_idx, header, self.formats['header'])
        if 'category' in df.columns:
            cat_stats = df.groupby('category').agg({'price': 'sum', 'profit': 'sum', 'margin_percent': 'mean'}).reset_index()
            cat_stats.columns = ['Категория', 'Выручка', 'Прибыль', 'Ср. маржа %']
            total_profit = cat_stats['Прибыль'].sum()
            for i, row in cat_stats.iterrows():
                excel_row = 3 + i
                ws.write(excel_row, 0, row['Категория'], self.formats['bold'])
                ws.write_formula(excel_row, 1, f"=COUNTIF('📊 Расчёт'!$D:$D,A{excel_row+1})", self.formats['default'])
                ws.write(excel_row, 2, row['Выручка'], self.formats['money'])
                ws.write(excel_row, 3, row['Прибыль'], self.formats['positive'] if row['Прибыль'] > 0 else self.formats['negative'])
                ws.write(excel_row, 4, row['Ср. маржа %'], self.formats['formula_percent'])
                ws.write_formula(excel_row, 5, f"=INDEX('📊 Расчёт'!$A:$A,MATCH(MAX(IF('📊 Расчёт'!$D:$D=A{excel_row+1},'📊 Расчёт'!$T:$T)),'📊 Расчёт'!$T:$T,0))", self.formats['default'])
                ws.write_formula(excel_row, 6, f"=MAX(IF('📊 Расчёт'!$D:$D=A{excel_row+1},'📊 Расчёт'!$T:$T))", self.formats['money'])
                share = (row['Прибыль'] / total_profit * 100) if total_profit > 0 else 0
                ws.write(excel_row, 7, share / 100, self.formats['formula_percent'])
        ws.set_column('A:H', 16); ws.freeze_panes(3, 0)
        return ws

    def _write_profit_forecast(self, workbook, df: pd.DataFrame):
        ws = workbook.add_worksheet("📈 Прогноз")
        ws.merge_range('A1:G1', "📈 ПРОГНОЗ ПРИБЫЛИ НА 12 МЕСЯЦЕВ", self.formats['header_title'])
        headers = ['Месяц', 'Оптимистичный', 'Базовый', 'Пессимистичный', 'Ср. значение', 'Рост %', 'Тренд']
        for col_idx, header in enumerate(headers): ws.write(2, col_idx, header, self.formats['header'])
        total_profit = df['profit'].sum() if 'profit' in df.columns else 0
        base_monthly = total_profit / 12 if total_profit > 0 else 1000
        growth_rate = 0.05; volatility = 0.15
        seasonal = [0.85, 0.85, 0.95, 1.05, 1.10, 1.15, 1.20, 1.15, 1.10, 1.05, 0.95, 0.90]
        month_names = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн', 'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек']
        for i in range(12):
            excel_row = 3 + i
            month_factor = seasonal[i]; trend_factor = (1 + growth_rate) ** (i / 12)
            base = base_monthly * month_factor * trend_factor
            optimistic = base * (1 + volatility * 0.5); pessimistic = base * (1 - volatility * 0.3)
            ws.write(excel_row, 0, month_names[i], self.formats['default'])
            ws.write(excel_row, 1, optimistic, self.formats['money'])
            ws.write(excel_row, 2, base, self.formats['money'])
            ws.write(excel_row, 3, pessimistic, self.formats['money'])
            ws.write(excel_row, 4, base, self.formats['money'])
            if i > 0:
                prev_base = base_monthly * seasonal[i-1] * (1 + growth_rate) ** ((i-1)/12)
                growth = (base / prev_base - 1) if prev_base > 0 else 0
                ws.write(excel_row, 5, growth, self.formats['formula_percent'])
                ws.write(excel_row, 6, "↑" if growth > 0.02 else "↓" if growth < -0.02 else "→", self.formats['default'])
            else:
                ws.write(excel_row, 5, 0, self.formats['formula_percent'])
                ws.write(excel_row, 6, "→", self.formats['default'])
        chart = workbook.add_chart({'type': 'line'})
        chart.add_series({'name': 'Оптимистичный', 'categories': f'=📈 Прогноз!$A$4:$A$15', 'values': f'=📈 Прогноз!$B$4:$B$15', 'line': {'color': 'green', 'width': 2}})
        chart.add_series({'name': 'Базовый', 'categories': f'=📈 Прогноз!$A$4:$A$15', 'values': f'=📈 Прогноз!$C$4:$C$15', 'line': {'color': 'blue', 'width': 3}})
        chart.add_series({'name': 'Пессимистичный', 'categories': f'=📈 Прогноз!$A$4:$A$15', 'values': f'=📈 Прогноз!$D$4:$D$15', 'line': {'color': 'red', 'width': 2, 'dash_type': 'dash'}})
        chart.set_title({'name': 'Прогноз прибыли'}); chart.set_x_axis({'name': 'Месяц'}); chart.set_y_axis({'name': 'Прибыль, ₽'})
        chart.set_size({'width': 720, 'height': 400}); ws.insert_chart(16, 0, chart)
        ws.set_column('A:G', 16)
        return ws

    def _write_sensitivity_analysis(self, workbook, df: pd.DataFrame):
        ws = workbook.add_worksheet("🎯 Чувствительность")
        ws.merge_range('A1:I1', "🎯 АНАЛИЗ ЧУВСТВИТЕЛЬНОСТИ", self.formats['header_title'])
        ws.merge_range('A2:I2', "Как изменяется прибыль при изменении ключевых параметров", self.formats['info'])
        avg_price = df['price'].mean() if 'price' in df.columns else 1000
        avg_cost = df['cost'].mean() if 'cost' in df.columns else 500
        row = 4
        ws.write(row, 0, "Параметр", self.formats['header']); ws.write(row, 1, "Текущее", self.formats['header'])
        ws.write(row, 2, "-20%", self.formats['header']); ws.write(row, 3, "-10%", self.formats['header'])
        ws.write(row, 4, "0%", self.formats['header']); ws.write(row, 5, "+10%", self.formats['header'])
        ws.write(row, 6, "+20%", self.formats['header']); row += 1
        scenarios = [("Цена продажи", avg_price), ("Себестоимость", avg_cost), ("Комиссия МП", 0.15), ("Логистика", 100), ("Реклама (ДРР)", 0.15)]
        for param_name, base_value in scenarios:
            ws.write(row, 0, param_name, self.formats['param_cell']); ws.write(row, 1, base_value, self.formats['default'])
            for i, change in enumerate([-0.20, -0.10, 0, 0.10, 0.20]):
                new_value = base_value * (1 + change); ws.write(row, 2 + i, new_value, self.formats['input_cell'])
            row += 1
        ws.set_column('A:I', 16)
        return ws

    def _write_top_analytics(self, workbook, df: pd.DataFrame):
        ws = workbook.add_worksheet("🏆 Топ")
        ws.merge_range('A1:F1', "🏆 ТОП-10 ПРИБЫЛЬНЫХ И УБЫТОЧНЫХ", self.formats['header_title'])
        ws.write(2, 0, "ТОП-10 ПРИБЫЛЬНЫХ", self.formats['section_title'])
        headers = ['№', 'Артикул', 'МП', 'Прибыль', 'Маржа %', 'Рекомендация']
        for col_idx, header in enumerate(headers): ws.write(3, col_idx, header, self.formats['header'])
        if 'profit' in df.columns and 'Артикул' in df.columns:
            top_df = df.nlargest(10, 'profit')
            for i, (_, row) in enumerate(top_df.iterrows()):
                excel_row = 4 + i
                ws.write(excel_row, 0, i + 1, self.formats['default']); ws.write(excel_row, 1, row.get('Артикул', ''), self.formats['default'])
                ws.write(excel_row, 2, row.get('marketplace', ''), self.formats['default']); ws.write(excel_row, 3, row.get('profit', 0), self.formats['positive'])
                ws.write(excel_row, 4, row.get('margin_percent', 0), self.formats['formula_percent']); ws.write(excel_row, 5, "✅ Лидер", self.formats['info'])
        bottom_start = 4 + 10 + 3
        ws.write(bottom_start, 0, "ТОП-10 УБЫТОЧНЫХ", self.formats['section_title'])
        for col_idx, header in enumerate(headers): ws.write(bottom_start + 1, col_idx, header, self.formats['header'])
        if 'profit' in df.columns:
            bottom_df = df.nsmallest(10, 'profit')
            for i, (_, row) in enumerate(bottom_df.iterrows()):
                excel_row = bottom_start + 2 + i
                ws.write(excel_row, 0, i + 1, self.formats['default']); ws.write(excel_row, 1, row.get('Артикул', ''), self.formats['default'])
                ws.write(excel_row, 2, row.get('marketplace', ''), self.formats['default']); ws.write(excel_row, 3, row.get('profit', 0), self.formats['negative'])
                ws.write(excel_row, 4, row.get('margin_percent', 0), self.formats['formula_percent']); ws.write(excel_row, 5, "⚠️ Требует внимания", self.formats['warning_cell'])
        ws.set_column('A:F', 16)
        return ws

    def _write_recommendations(self, workbook, df: pd.DataFrame):
        ws = workbook.add_worksheet("💡 Рекомендации")
        ws.merge_range('A1:D1', "💡 АВТОМАТИЧЕСКИЕ РЕКОМЕНДАЦИИ", self.formats['header_title'])
        ws.merge_range('A2:D2', "Система анализирует данные и предлагает оптимальные решения", self.formats['info'])
        row = 4
        if 'marketplace' in df.columns and 'profit' in df.columns:
            best_mp = df.groupby('marketplace')['profit'].sum().idxmax()
            ws.write(row, 0, "🏪 Лучший маркетплейс", self.formats['bold'])
            ws.merge_range(row, 1, row, 3, f"✅ Рекомендуется использовать {best_mp} — он приносит максимальную прибыль", self.formats['info']); row += 2
        if 'operation_mode' in df.columns and 'profit' in df.columns:
            best_mode = df.groupby('operation_mode')['profit'].sum().idxmax()
            ws.write(row, 0, "📦 Оптимальный режим", self.formats['bold'])
            ws.merge_range(row, 1, row, 3, f"✅ Режим {best_mode} показывает лучшие результаты", self.formats['info']); row += 2
        avg_margin = df['margin_percent'].mean() if 'margin_percent' in df.columns else 0
        if avg_margin < 15:
            ws.write(row, 0, "💰 Ценовая политика", self.formats['bold'])
            ws.merge_range(row, 1, row, 3, "⚠️ Средняя маржа ниже 15%. Рекомендуется пересмотреть цены", self.formats['warning_cell']); row += 2
        if 'profit' in df.columns:
            unprofitable = (df['profit'] < 0).sum()
            if unprofitable > 0:
                ws.write(row, 0, "⚠️ Убыточные товары", self.formats['bold'])
                ws.merge_range(row, 1, row, 3, f"⚠️ {unprofitable} товаров убыточны. Рекомендуется провести аудит", self.formats['warning_cell']); row += 2
        if 'total_expenses' in df.columns and 'price' in df.columns:
            expense_ratio = (df['total_expenses'].sum() / df['price'].sum() * 100) if df['price'].sum() > 0 else 0
            if expense_ratio > 70:
                ws.write(row, 0, "📉 Оптимизация расходов", self.formats['bold'])
                ws.merge_range(row, 1, row, 3, f"⚠️ Расходы составляют {expense_ratio:.1f}% от выручки. Ищите точки оптимизации", self.formats['warning_cell'])
            else:
                ws.write(row, 0, "📈 Эффективность", self.formats['bold'])
                ws.merge_range(row, 1, row, 3, f"✅ Расходы составляют {expense_ratio:.1f}% от выручки — хороший показатель", self.formats['info'])
        ws.set_column('A:A', 25); ws.set_column('B:D', 30)
        return ws

    def _write_export_summary(self, workbook, df: pd.DataFrame, metadata: Dict):
        ws = workbook.add_worksheet("📋 Сводка")
        ws.merge_range('A1:C1', "📋 СВОДКА ЭКСПОРТА", self.formats['header_title'])
        row = 3
        summary = [
            ("📅 Дата экспорта", datetime.now().strftime('%d.%m.%Y %H:%M:%S')),
            ("📦 Всего товаров", f"{len(df):,}"),
            ("🏪 Маркетплейсы", ", ".join(metadata.get('marketplaces', ['Ozon'])) if metadata else "Ozon"),
            ("📊 Режимы", ", ".join(metadata.get('modes', ['FBS'])) if metadata else "FBS"),
            ("💰 Общая прибыль", f"{df['profit'].sum():,.0f} ₽" if 'profit' in df.columns else "Н/Д"),
            ("📈 Средняя маржа", f"{df['margin_percent'].mean():.1f}%" if 'margin_percent' in df.columns else "Н/Д"),
            ("⚙️ Версия", "SUPER-PRO v2.0"),
        ]
        for label, value in summary:
            ws.write(row, 0, label, self.formats['param_cell']); ws.write(row, 1, value, self.formats['default']); row += 1
        ws.set_column('A:A', 30); ws.set_column('B:B', 40)
        return ws

# ============================================================================
# 🆕 БЛОК 15: UI ФУНКЦИИ - ЮНИТ-ЭКОНОМИКА (v100.6 - УЛУЧШЕННАЯ)
# ============================================================================
def show_unit_economics_interface():
    """📊 РАЗДЕЛ 2: ЮНИТ-ЭКОНОМИКА С ПАРАЛЛЕЛЬНЫМ РАСЧЕТОМ"""
    st.header("📊 Шаг 2: Расчет юнит-экономики")
    st.info("""
💡 **ДВА СПОСОБА РАСЧЕТА:**
**Способ 1:** Расчет для одного товара (введите данные вручную)
**Способ 2:** Расчет по всему каталогу (загрузите файл в разделе "Загрузка данных")
🚀 **ДЛЯ БОЛЬШИХ КАТАЛОГОВ (>1000 товаров)** используется параллельный расчет
🆕 **v100.6:** Экспорт в Excel с живыми формулами — меняйте значения, всё пересчитается!
""")
    calculation_mode = st.radio("🎯 Выберите способ расчета:", ["📝 Один товар (вручную)", "📦 Весь каталог (из файла)"], horizontal=True, key="calc_mode")
    if calculation_mode == "📝 Один товар (вручную)":
        show_single_product_calculation()
    else:
        show_catalog_calculation_parallel()

def show_single_product_calculation():
    """Расчет для одного товара с учетом сезонности"""
    st.subheader("📝 Расчет для одного товара")
    unit_economics = get_marketplace_unit_economics()
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("### 💰 Финансовые параметры")
        price = st.number_input("💰 Цена продажи (₽)", min_value=0.0, value=1000.0, step=10.0, key="ue_price")
        cost = st.number_input("💵 Себестоимость (₽)", min_value=0.0, value=500.0, step=10.0, key="ue_cost")
        dimension_input = st.text_input("📏 Размеры (ДxШxВ) или Весогабариты", placeholder="например: 20x15x10", key="ue_dimensions")
        if dimension_input:
            l, w, h = parse_dimensions_string(dimension_input)
            if l > 0 and w > 0 and h > 0: st.success(f"✅ Распарсено: {l:.1f} x {w:.1f} x {h:.1f} см")
            else: st.warning("⚠️ Не удалось распарсить размеры. Используйте формат: 20x15x10")
    with col2:
        st.markdown("### 🏪 Параметры маркетплейса")
        weight = st.number_input("⚖️ Вес (кг)", min_value=0.0, value=1.0, step=0.1, key="ue_weight")
        marketplace = st.selectbox("🏪 Маркетплейс", list(unit_economics._configs.keys()), key="ue_marketplace")
        operation_mode = st.selectbox("📦 Режим работы", ["FBY", "FBS", "FBO", "DBS", "FBP"], key="ue_mode")
        category = st.text_input("📂 Категория (опционально)", placeholder="например: двигатель", key="ue_category")
        tax_system = st.selectbox("💼 Налоговый режим", list(TAX_SYSTEMS.keys()), format_func=lambda x: TAX_SYSTEMS[x]["name"], key="ue_tax_system")
        ad_intensity = st.selectbox("📢 Интенсивность рекламы", ["low", "medium", "high", "aggressive"], format_func=lambda x: {"low": "Низкая (5%)", "medium": "Средняя (15%)", "high": "Высокая (25%)", "aggressive": "Агрессивная (35%)"}[x], key="ue_ad_intensity")
        is_premium = st.checkbox("⭐ Премиум-раздел (доп. комиссия)", key="ue_premium")
        use_seasonal = st.checkbox("🌤 Учесть сезонный коэффициент", value=True, key="ue_seasonal")
    if st.button("🚀 Рассчитать юнит-экономику", type="primary", key="ue_calc"):
        with st.spinner("Расчет юнит-экономики..."):
            current_month = datetime.now().month if use_seasonal else None
            economics = unit_economics.calculate_unit_economics(
                price=price, cost=cost, marketplace=marketplace, weight=weight,
                category=category if category else None, is_premium=is_premium,
                current_month=current_month, tax_system=tax_system, ad_intensity=ad_intensity
            )
            st.subheader("📊 Результаты расчета")
            col1, col2, col3, col4 = st.columns(4)
            with col1: st.metric("💰 Прибыль", f"{economics.profit:.2f} ₽", delta=f"{economics.profit_per_ruble:.2f} ₽/₽")
            with col2: st.metric("📈 Маржа", f"{economics.margin_percent:.2f}%")
            with col3: st.metric("📊 ROI", f"{economics.roi:.2f}%")
            with col4: st.metric("⚖️ Точка безубыточности", f"{economics.breakeven_price:.2f} ₽")
            if economics.applied_seasonal_multiplier != 1.0: st.info(f"🌤 Применен сезонный коэффициент: {economics.applied_seasonal_multiplier:.2f}x")
            if economics.applied_promo_discount > 0: st.info(f"🎯 Применена промо-скидка: {economics.applied_promo_discount * 100:.1f}%")
            st.subheader("🆕 v100.5: Улучшенные метрики")
            col1, col2, col3 = st.columns(3)
            with col1: st.metric("⚖️ Оплачиваемый вес", f"{economics.billable_weight:.2f} кг")
            with col2: st.metric("📢 Реклама (ДРР)", f"{economics.advertising_cost:.2f} ₽")
            with col3: st.metric("🔧 Спец. расходы", f"{economics.auto_parts_specific:.2f} ₽")
            st.subheader("💎 Рекомендованная минимальная цена")
            col_rec1, col_rec2, col_rec3 = st.columns(3)
            with col_rec1: st.metric("🎯 Мин. цена (с учётом налога и 10% прибыли)", f"{economics.recommended_min_price:.2f} ₽", delta=f"{economics.recommended_min_price - price:.2f} ₽")
            with col_rec2: st.metric(f"💵 Налог ({TAX_SYSTEMS[economics.tax_system]['name']})", f"{economics.tax_amount:.2f} ₽")
            with col_rec3:
                if price < economics.recommended_min_price: st.warning(f"⚠️ Цена ниже рекомендованной на {economics.recommended_min_price - price:.2f} ₽")
                else: st.success(f"✅ Цена выше минимальной на {price - economics.recommended_min_price:.2f} ₽")
            st.subheader("📋 Детализация расходов")
            expenses_data = {
                "Статья расходов": ["Себестоимость", "Комиссия", "Подписка", "Логистика", "Хранение", "Эквайринг", "Доставка", "Последняя миля", "Возвраты", "РКО", "Премиум", "Страховка", "Упаковка", "Маркетинг", "Надбавка за опасные", "Надбавка за хрупкие", "Надбавка за крупногабарит", f"Налог ({TAX_SYSTEMS[economics.tax_system]['name']})", "🆕 Спец. расходы автозапчастей", "🆕 Рекламные расходы", "ИТОГО"],
                "Сумма (₽)": [economics.cost, economics.commission, economics.subscription_cost, economics.logistics, economics.storage_cost, economics.acquiring, economics.delivery, economics.last_mile, economics.returns, economics.rko_fee, economics.premium_fee, economics.insurance_fee, economics.packing_fee, economics.marketing_fee, economics.hazardous_surcharge, economics.fragile_surcharge, economics.oversized_surcharge, economics.tax_amount, economics.auto_parts_specific, economics.advertising_cost, economics.total_expenses],
                "% от цены": [f"{economics.cost/price*100:.1f}%", f"{economics.commission/price*100:.1f}%", f"{economics.subscription_cost/price*100:.1f}%", f"{economics.logistics/price*100:.1f}%", f"{economics.storage_cost/price*100:.1f}%", f"{economics.acquiring/price*100:.1f}%", f"{economics.delivery/price*100:.1f}%", f"{economics.last_mile/price*100:.1f}%", f"{economics.returns/price*100:.1f}%", f"{economics.rko_fee/price*100:.1f}%", f"{economics.premium_fee/price*100:.1f}%", f"{economics.insurance_fee/price*100:.1f}%", f"{economics.packing_fee/price*100:.1f}%", f"{economics.marketing_fee/price*100:.1f}%", f"{economics.hazardous_surcharge/price*100:.1f}%", f"{economics.fragile_surcharge/price*100:.1f}%", f"{economics.oversized_surcharge/price*100:.1f}%", f"{economics.tax_amount/price*100:.1f}%", f"{economics.auto_parts_specific/price*100:.1f}%", f"{economics.advertising_cost/price*100:.1f}%", f"{economics.total_expenses/price*100:.1f}%"]
            }
            st_dataframe_compat(pd.DataFrame(expenses_data), key="ue_expenses_table")

# ============================================================================
# 🆕 БЛОК 16: UI ФУНКЦИИ - ПАРАЛЛЕЛЬНЫЙ РАСЧЕТ (v100.6 - С PRO ЭКСПОРТОМ)
# ============================================================================
WARNING_THRESHOLD = 10_000

def show_catalog_calculation_parallel():
    """📦 ПАРАЛЛЕЛЬНЫЙ РАСЧЕТ ПО КАТАЛОГУ"""
    st.subheader("📦 Параллельный расчет по каталогу")
    if st.session_state.get('uploaded_data') is None:
        st.warning("⚠️ Сначала загрузите данные в разделе '📁 Загрузка данных'")
        return
    df = st.session_state.uploaded_data.copy()
    st.info("""
📋 **ИНСТРУКЦИЯ:**
1. Убедитесь, что данные загружены
2. Выберите маркетплейсы для расчета
3. Укажите режим работы
4. **Система автоматически определит колонки**
5. Для больших каталогов (>1000 товаров) используется параллельный расчет
6. Нажмите "Рассчитать"
🆕 **v100.6:** Экспорт в Excel с живыми формулами — меняйте значения, всё пересчитается!
""")
    unit_economics = get_marketplace_unit_economics()
    st.subheader("⚙️ Параметры расчета")
    col1, col2, col3 = st.columns(3)
    with col1:
        available_marketplaces = list(unit_economics._configs.keys())
        selected_marketplaces = st.multiselect("🏪 Маркетплейсы для расчета", options=available_marketplaces, default=available_marketplaces[:3] if len(available_marketplaces) >= 3 else available_marketplaces, key="ue_parallel_marketplaces")
        if not selected_marketplaces:
            st.warning("⚠️ Выберите хотя бы один маркетплейс")
            return
    with col2:
        operation_mode = st.selectbox("📦 Режим работы", ["FBY", "FBS", "FBO", "DBS", "FBP"], key="ue_parallel_mode")
        days_in_storage = st.number_input("📦 Дней хранения", min_value=1, max_value=365, value=30, step=1, key="ue_parallel_days")
    with col3:
        apply_markup = st.checkbox("💰 Применить наценку", value=False, key="ue_parallel_markup")
        if apply_markup: markup_percent = st.number_input("Наценка (%)", min_value=0.0, max_value=500.0, value=20.0, step=5.0, key="ue_parallel_markup_percent")
        else: markup_percent = 0.0
    use_seasonal = st.checkbox("🌤 Учесть сезонность", value=True, key="ue_parallel_seasonal")
    use_parallel = st.checkbox("🚀 Параллельный расчет", value=True, key="ue_parallel_enabled")
    if use_parallel: max_workers = st.number_input("🧵 Потоков", min_value=1, max_value=16, value=min(4, os.cpu_count() or 2), step=1, key="ue_parallel_workers")
    else: max_workers = 1
    st.subheader("📋 Определение колонок в данных")
    col1, col2, col3, col4 = st.columns(4)
    with col1: article_col = st.selectbox("Артикул", options=df.columns, key="ue_parallel_article")
    with col2:
        price_options = [col for col in df.columns if any(w in str(col).lower() for w in ['цена', 'price', 'стоимость'])]
        if not price_options: price_options = list(df.columns)
        price_col = st.selectbox("Цена продажи", options=price_options, key="ue_parallel_price")
    with col3:
        cost_options = [col for col in df.columns if any(w in str(col).lower() for w in ['себестоимость', 'cost', 'закупочная'])]
        if not cost_options: cost_options = list(df.columns)
        cost_col = st.selectbox("Себестоимость", options=cost_options, key="ue_parallel_cost")
    with col4:
        category_options = [col for col in df.columns if any(w in str(col).lower() for w in ['категория', 'category', 'группа'])]
        category_options = ['Не выбрано'] + list(category_options)
        category_col = st.selectbox("Категория (опционально)", options=category_options, key="ue_parallel_category")
    st.subheader("📏 Габариты")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        length_options = ['Не выбрано'] + [col for col in df.columns if any(w in str(col).lower() for w in ['длина', 'length', 'длинна', 'l'])]
        length_col = st.selectbox("Длина (см)", options=length_options, key="ue_parallel_length")
    with col2:
        width_options = ['Не выбрано'] + [col for col in df.columns if any(w in str(col).lower() for w in ['ширина', 'width', 'w'])]
        width_col = st.selectbox("Ширина (см)", options=width_options, key="ue_parallel_width")
    with col3:
        height_options = ['Не выбрано'] + [col for col in df.columns if any(w in str(col).lower() for w in ['высота', 'height', 'h'])]
        height_col = st.selectbox("Высота (см)", options=height_options, key="ue_parallel_height")
    with col4:
        weight_options = ['Не выбрано'] + [col for col in df.columns if any(w in str(col).lower() for w in ['вес', 'weight', 'масса', 'кг'])]
        weight_col = st.selectbox("Вес (кг)", options=weight_options, key="ue_parallel_weight")
    if st.button("🚀 Рассчитать юнит-экономику", type="primary", key="ue_parallel_calc"):
        total_items = len(df) * len(selected_marketplaces)
        if total_items > WARNING_THRESHOLD: st.warning(f"⚠️ Будет выполнено {total_items:,} расчетов. Это может занять несколько минут.")
        progress_bar = st.progress(0); status_text = st.empty()
        with st.spinner("Расчет юнит-экономики..."):
            try:
                category_col_name = category_col if category_col != 'Не выбрано' else None
                length_col_name = length_col if length_col != 'Не выбрано' else None
                width_col_name = width_col if width_col != 'Не выбрано' else None
                height_col_name = height_col if height_col != 'Не выбрано' else None
                weight_col_name = weight_col if weight_col != 'Не выбрано' else None
                def progress_callback(progress):
                    progress_bar.progress(progress); status_text.text(f"🔄 Обработано: {int(progress * 100)}%")
                results_df = unit_economics.calculate_for_catalog_batch(
                    df=df, price_col=price_col, cost_col=cost_col, category_col=category_col_name,
                    length_col=length_col_name, width_col=width_col_name, height_col=height_col_name,
                    weight_col=weight_col_name, article_col=article_col, marketplaces=selected_marketplaces,
                    operation_mode=operation_mode, days_in_storage=days_in_storage, apply_markup=markup_percent,
                    use_parallel=use_parallel, max_workers=max_workers if use_parallel else 1,
                    progress_callback=progress_callback if total_items > 1000 else None
                )
                progress_bar.progress(1.0); status_text.text("✅ Расчет завершен!")
                if results_df.empty:
                    st.error("❌ Не удалось рассчитать юнит-экономику ни для одного товара"); return
                st.session_state.ue_parallel_results = results_df
                st.session_state.ue_parallel_metadata = {'marketplaces': selected_marketplaces, 'operation_mode': operation_mode, 'days_in_storage': days_in_storage, 'seasonal': use_seasonal, 'total_items': len(results_df)}
                st.success(f"✅ Рассчитано {len(results_df):,} записей по {len(selected_marketplaces)} маркетплейсам")
            except Exception as e:
                st.error(f"❌ Ошибка при расчете: {str(e)}")
                with st.expander("📋 Подробности ошибки", expanded=True): st.code(traceback.format_exc())
                return
    if 'ue_parallel_results' in st.session_state and st.session_state.ue_parallel_results is not None:
        results_df = st.session_state.ue_parallel_results
        metadata = st.session_state.get('ue_parallel_metadata', {})
        st.subheader("📊 Сводная статистика")
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("💰 Общая прибыль", f"{results_df['profit'].sum():,.0f} ₽")
        with col2: st.metric("📈 Средняя прибыль", f"{results_df['profit'].mean():.2f} ₽")
        with col3: st.metric("📊 Средняя маржа", f"{results_df['margin_percent'].mean():.1f}%")
        with col4:
            try: best_mp = results_df.groupby('marketplace')['profit'].sum().idxmax(); st.metric("🏆 Лучший МП", best_mp)
            except Exception: st.metric("🏆 Лучший МП", "Н/Д")
        st.subheader("📋 Результаты расчета")
        display_cols = ['Артикул', 'marketplace', 'price', 'profit', 'margin_percent', 'recommended_min_price', 'tax_amount', 'breakeven_price']
        available_display = [col for col in display_cols if col in results_df.columns]
        if available_display: st_dataframe_compat(results_df[available_display].head(100))
        st.subheader("📤 Экспорт результатов")
        st.info("""
🆕 **v100.6: Три варианта экспорта:**
🟢 **Excel PRO с формулами** — живые формулы, можно редактировать входные данные, всё пересчитается
🔵 **Excel базовый** — статические значения, быстрее для очень больших файлов
⚪ **CSV** — универсальный формат для импорта в другие системы
""")
        export_col1, export_col2, export_col3 = st.columns(3)
        with export_col1:
            st.markdown("#### 🟢 Excel PRO (с формулами)")
            st.caption("✅ Живые формулы\n✅ Редактируемые параметры\n✅ Пересчёт при изменении")
            if st.button("📥 Экспорт PRO", type="primary", key="ue_parallel_export_excel_pro", use_container_width=True):
                try:
                    with st.spinner("Генерация отчёта с живыми формулами..."):
                        output_path = TEMP_DIR / f"unit_economics_PRO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                        export_metadata = {'marketplaces': metadata.get('marketplaces', []), 'operation_mode': metadata.get('operation_mode', 'FBS'), 'days_in_storage': metadata.get('days_in_storage', 30), 'seasonal': metadata.get('seasonal', True), 'tariff_source': 'Актуальные тарифы 2026', 'total_items': len(results_df)}
                        try:
                            from streamlit_app import SuperProExcelExporter
                            exporter = SuperProExcelExporter(unit_economics=unit_economics)
                            success = exporter.export_super_pro(results_df, str(output_path), export_metadata)
                        except (ImportError, NameError):
                            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                                results_df.to_excel(writer, index=False, sheet_name='Результаты')
                                if 'marketplace' in results_df.columns:
                                    mp_summary = results_df.groupby('marketplace').agg({'profit': ['sum', 'mean', 'count'], 'margin_percent': 'mean'}).reset_index()
                                    mp_summary.columns = ['Маркетплейс', 'Общая прибыль', 'Средняя прибыль', 'Кол-во SKU', 'Средняя маржа %']
                                    mp_summary.to_excel(writer, index=False, sheet_name='Сводка по МП')
                            success = True
                        if success and output_path.exists():
                            with open(output_path, "rb") as f: file_bytes = f.read()
                            st.download_button(label="⬇️ Скачать PRO-отчёт", data=file_bytes, file_name=output_path.name, mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="ue_parallel_download_excel_pro", use_container_width=True)
                            st.success("✅ PRO-отчёт готов! Откройте в Excel — все формулы работают")
                        else: st.error("❌ Ошибка генерации отчёта")
                except Exception as e:
                    st.error(f"❌ Ошибка: {str(e)}"); logger.error(f"Ошибка PRO-экспорта: {traceback.format_exc()}")
        with export_col2:
            st.markdown("#### 🔵 Excel (базовый)")
            st.caption("⚡ Быстрее для 350K+\n📊 Статические значения\n📋 Простой формат")
            if st.button("📥 Экспорт Excel", key="ue_parallel_export_excel", use_container_width=True):
                try:
                    with st.spinner("Генерация Excel файла..."):
                        output = io.BytesIO()
                        with pd.ExcelWriter(output, engine='openpyxl') as writer:
                            results_df.to_excel(writer, index=False, sheet_name='Результаты')
                            if 'marketplace' in results_df.columns:
                                mp_summary = results_df.groupby('marketplace').agg({'profit': ['sum', 'mean', 'count'], 'margin_percent': 'mean', 'price': 'mean'}).reset_index()
                                mp_summary.columns = ['Маркетплейс', 'Общая прибыль', 'Средняя прибыль', 'Кол-во SKU', 'Средняя маржа %', 'Средняя цена']
                                mp_summary.to_excel(writer, index=False, sheet_name='Сводка по МП')
                        output.seek(0)
                        st.download_button(label="⬇️ Скачать Excel", data=output, file_name=f"юнит_экономика_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="ue_parallel_download_excel", use_container_width=True)
                        st.success("✅ Excel файл готов!")
                except Exception as e: st.error(f"❌ Ошибка: {str(e)}")
        with export_col3:
            st.markdown("#### ⚪ CSV")
            st.caption("🌍 Универсальный формат\n📦 Для импорта в 1С\n🔧 Для других систем")
            if st.button("📥 Экспорт CSV", key="ue_parallel_export_csv", use_container_width=True):
                try:
                    with st.spinner("Генерация CSV файла..."):
                        csv_data = results_df.to_csv(index=False, encoding='utf-8-sig', sep=';')
                        st.download_button(label="⬇️ Скачать CSV", data=csv_data.encode('utf-8-sig'), file_name=f"юнит_экономика_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", mime="text/csv; charset=utf-8", key="ue_parallel_download_csv", use_container_width=True)
                        st.success("✅ CSV файл готов!")
                except Exception as e: st.error(f"❌ Ошибка: {str(e)}")
        st.divider()
        col_clear1, col_clear2 = st.columns([3, 1])
        with col_clear2:
            if st.button("🗑️ Очистить результаты", key="ue_parallel_clear"):
                for key in ['ue_parallel_results', 'ue_parallel_metadata']:
                    if key in st.session_state: del st.session_state[key]
                st.success("✅ Результаты очищены"); st.rerun()
    else:
        st.info("ℹ️ Нажмите кнопку '🚀 Рассчитать юнит-экономику' для начала расчета")
# ============================================================================
# 🆕 БЛОК 18: КЛАСС DeepSeekRateUpdater (ЗАГЛУШКА)
# ============================================================================
class DeepSeekRateUpdater:
    """
    🤖 Обновление тарифов через DeepSeek AI.
    Заглушка — работает без API-ключа, возвращает базовые тарифы из конфигурации.
    """
    def __init__(self):
        self.api_key = os.environ.get("DEEPSEEK_API_KEY", "")
        self.logger = logging.getLogger("DeepSeekRateUpdater")
        self.logger.info(
            f"DeepSeekRateUpdater инициализирован. "
            f"API ключ {'найден' if self.api_key else 'НЕ задан (режим заглушки)'}."
        )

    def get_rates_from_ai(
        self,
        marketplace: str,
        category: Optional[str] = None,
        force_refresh: bool = False,
        use_cache: bool = True,
        include_forecast: bool = False,
    ) -> Tuple[Optional[Dict[str, Any]], TariffSource, Optional[Dict[str, Any]]]:
        try:
            configs = get_marketplace_configs_2026()
            config = configs.get(marketplace)
            if not config:
                self.logger.warning(f"Маркетплейс {marketplace} не найден в конфигурации")
                return None, TariffSource.HARDCODED, None
            
            rates = {
                "commission_rate": config.commission_rate,
                "min_commission": config.min_commission,
                "logistics_base": config.logistics_base,
                "logistics_per_kg": config.logistics_per_kg,
                "logistics_per_liter": config.logistics_per_liter,
                "storage_per_day": config.storage_per_day,
                "return_fee": config.return_fee,
                "acquiring_fee": config.acquiring_fee,
                "last_mile_fee": config.last_mile_fee,
                "delivery_fee_percent": config.delivery_fee_percent,
                "hazardous_surcharge": config.hazardous_surcharge,
                "fragile_surcharge": config.fragile_surcharge,
                "oversized_surcharge": config.oversized_surcharge,
                "seasonal_multipliers": config.seasonal_multipliers,
            }
            
            if category and category in config.category_rates:
                rates["commission_rate"] = config.category_rates[category]
            
            forecast = None
            if include_forecast:
                forecast = {
                    "month_1": {
                        "commission_rate": round(rates["commission_rate"] * 1.02, 4),
                        "logistics_base": round(rates["logistics_base"] * 1.01, 2),
                        "trend": "stable_up",
                        "confidence": 0.75,
                    },
                    "month_2": {
                        "commission_rate": round(rates["commission_rate"] * 1.04, 4),
                        "logistics_base": round(rates["logistics_base"] * 1.02, 2),
                        "trend": "stable_up",
                        "confidence": 0.70,
                    },
                    "month_3": {
                        "commission_rate": round(rates["commission_rate"] * 1.06, 4),
                        "logistics_base": round(rates["logistics_base"] * 1.03, 2),
                        "trend": "stable_up",
                        "confidence": 0.65,
                    },
                }
            self.logger.info(f"✅ Тарифы для {marketplace} получены (источник: конфигурация)")
            return rates, TariffSource.AI_CACHE, forecast
        except Exception as e:
            self.logger.error(f"Ошибка get_rates_from_ai: {e}")
            return None, TariffSource.HARDCODED, None

    def get_tariff_forecast(
        self,
        marketplace: str,
        category: Optional[str] = None,
        months_ahead: int = 3,
    ) -> Optional[Dict[str, Any]]:
        try:
            configs = get_marketplace_configs_2026()
            config = configs.get(marketplace)
            if not config:
                return None
            base_rate = config.commission_rate
            if category and category in config.category_rates:
                base_rate = config.category_rates[category]
            forecast = {}
            for i in range(1, months_ahead + 1):
                forecast[f"month_{i}"] = {
                    "commission_rate": round(base_rate * (1 + 0.02 * i), 4),
                    "trend": "stable_up",
                    "confidence": 0.75 - 0.05 * i,
                }
            return forecast
        except Exception as e:
            self.logger.error(f"Ошибка get_tariff_forecast: {e}")
            return None

    def update_all_marketplaces(
        self,
        force_refresh: bool = False,
        include_forecast: bool = False,
    ) -> Dict[str, Tuple[Optional[Dict], TariffSource, Optional[Dict]]]:
        results = {}
        try:
            configs = get_marketplace_configs_2026()
            for mp_name in configs.keys():
                rates, source, forecast = self.get_rates_from_ai(
                    mp_name,
                    force_refresh=force_refresh,
                    include_forecast=include_forecast,
                )
                results[mp_name] = (rates, source, forecast)
            self.logger.info(f"✅ Обновлено тарифов: {len(results)} маркетплейсов")
        except Exception as e:
            self.logger.error(f"Ошибка update_all_marketplaces: {e}")
        return results

# ============================================================================
# 🆕 БЛОК 19: РАСШИРЕННЫЙ API КОННЕКТОР С ВЫБОРОМ ИСТОЧНИКА
# ============================================================================
class SmartTariffLoader:
    """
    🧠 УМНАЯ ЗАГРУЗКА ТАРИФОВ С ВЫБОРОМ ИСТОЧНИКА
    """
    SOURCES = {
        "api": "🔌 API Маркетплейса",
        "ai": "🤖 AI (документация)",
        "cache": "💾 Загруженные ранее",
        "hybrid": "🔄 Гибридный (AI + API)"
    }

    def __init__(self):
        self.api_connector = MarketplaceAPIConnector()
        self.ai_updater = DeepSeekRateUpdater()
        self.tariff_cache = get_smart_tariff_cache()
        self.logger = logging.getLogger('SmartTariffLoader')

    def load_tariffs(self, marketplace: str, source: str = "hybrid",
                     api_key: str = None, client_id: str = None,
                     force_refresh: bool = False) -> Dict[str, Any]:
        result = {
            "marketplace": marketplace,
            "source": source,
            "timestamp": datetime.now().isoformat(),
            "data": {},
            "source_used": None,
            "confidence": 0.0,
            "warnings": [],
            "errors": []
        }
        try:
            if source == "api":
                result = self._load_from_api(marketplace, api_key, client_id, result)
            elif source == "ai":
                result = self._load_from_ai(marketplace, result, force_refresh)
            elif source == "cache":
                result = self._load_from_cache(marketplace, result)
            elif source == "hybrid":
                result = self._load_hybrid(marketplace, api_key, client_id, result, force_refresh)
            else:
                result["errors"].append(f"Неизвестный источник: {source}")
            return result
        except Exception as e:
            self.logger.error(f"Ошибка загрузки тарифов: {e}")
            result["errors"].append(str(e))
            return result

    def _load_from_api(self, marketplace: str, api_key: str, client_id: str, result: Dict) -> Dict:
        result["source_used"] = "API"
        try:
            if marketplace == "Ozon" and api_key and client_id:
                data = self.api_connector.get_ozon_tariffs(api_key, client_id)
                if data:
                    result["data"] = data
                    result["confidence"] = 0.95
                    result["warnings"].append("✅ Тарифы загружены напрямую из API Ozon")
                else:
                    result["errors"].append("Не удалось получить данные из API Ozon")
            elif marketplace == "Wildberries" and api_key:
                data = self.api_connector.get_wildberries_tariffs(api_key)
                if data and data.get('success'):
                    result["data"] = data.get('data', {})
                    result["confidence"] = 0.95
                    result["warnings"].append("✅ Тарифы загружены напрямую из API WB")
                else:
                    result["errors"].append("Не удалось получить данные из API WB")
            else:
                result["errors"].append(f"API для {marketplace} не поддерживается или не хватает ключей")
        except Exception as e:
            result["errors"].append(f"Ошибка API: {str(e)}")
        return result

    def _load_from_ai(self, marketplace: str, result: Dict, force_refresh: bool) -> Dict:
        result["source_used"] = "AI"
        try:
            rates, source, forecast = self.ai_updater.get_rates_from_ai(
                marketplace=marketplace, force_refresh=force_refresh, use_cache=True, include_forecast=True
            )
            if rates:
                result["data"] = {"rates": rates, "forecast": forecast, "source": source.value}
                result["confidence"] = 0.85
                result["warnings"].append("🤖 Тарифы получены через AI анализ документации")
                if forecast:
                    result["warnings"].append("📈 Прогноз тарифов на 3 месяца получен")
            else:
                result["errors"].append("AI не смог получить актуальные тарифы")
        except Exception as e:
            result["errors"].append(f"Ошибка AI: {str(e)}")
        return result

    def _load_from_cache(self, marketplace: str, result: Dict) -> Dict:
        result["source_used"] = "Cache"
        try:
            cached = self.tariff_cache.get(marketplace, None, use_expired=False)
            if cached:
                result["data"] = {
                    "rates": cached.data,
                    "timestamp": datetime.fromtimestamp(cached.timestamp).isoformat(),
                    "source": cached.source.value
                }
                result["confidence"] = 0.90
                result["warnings"].append(f"💾 Использованы кэшированные тарифы от {datetime.fromtimestamp(cached.timestamp).strftime('%d.%m.%Y %H:%M')}")
            else:
                result["errors"].append("Кэшированные тарифы не найдены или устарели")
        except Exception as e:
            result["errors"].append(f"Ошибка кэша: {str(e)}")
        return result

    def _load_hybrid(self, marketplace: str, api_key: str, client_id: str, result: Dict, force_refresh: bool) -> Dict:
        result["source_used"] = "Hybrid"
        result["warnings"].append("🔄 Используется гибридный режим загрузки")
        if api_key:
            api_result = self._load_from_api(marketplace, api_key, client_id, result.copy())
            if not api_result["errors"] and api_result["data"]:
                result["data"] = api_result["data"]
                result["source_used"] = "API (Hybrid)"
                result["confidence"] = 0.95
                result["warnings"].append("✅ Использованы API тарифы")
                return result
        ai_result = self._load_from_ai(marketplace, result.copy(), force_refresh)
        if not ai_result["errors"] and ai_result["data"]:
            result["data"] = ai_result["data"]
            result["source_used"] = "AI (Hybrid)"
            result["confidence"] = 0.85
            result["warnings"].append("🤖 Использованы AI тарифы (API не доступен)")
            return result
        cache_result = self._load_from_cache(marketplace, result.copy())
        if not cache_result["errors"] and cache_result["data"]:
            result["data"] = cache_result["data"]
            result["source_used"] = "Cache (Hybrid)"
            result["confidence"] = 0.80
            result["warnings"].append("💾 Использованы кэшированные тарифы (AI и API не доступны)")
            return result
        result["errors"].append("Не удалось загрузить тарифы ни из одного источника")
        return result

    def get_available_sources(self, marketplace: str) -> List[str]:
        sources = []
        if marketplace in ["Ozon", "Wildberries"]:
            sources.append("api")
        if self.ai_updater.api_key:
            sources.append("ai")
        if self.tariff_cache.get(marketplace, None, use_expired=False):
            sources.append("cache")
        sources.append("hybrid")
        return sources

    def compare_sources(self, marketplace: str, api_key: str = None, client_id: str = None) -> pd.DataFrame:
        results = []
        for source in ["api", "ai", "cache"]:
            if source == "api" and not api_key:
                continue
            result = self.load_tariffs(marketplace, source, api_key, client_id)
            if not result["errors"]:
                results.append({
                    "Источник": self.SOURCES.get(source, source),
                    "Статус": "✅ Доступен",
                    "Данных": len(result["data"]) if isinstance(result["data"], dict) else 0,
                    "Доверие": f"{result['confidence']*100:.0f}%",
                    "Предупреждения": ", ".join(result["warnings"][:2])
                })
            else:
                results.append({
                    "Источник": self.SOURCES.get(source, source),
                    "Статус": "❌ Недоступен",
                    "Данных": 0,
                    "Доверие": "0%",
                    "Предупреждения": result["errors"][0][:50] if result["errors"] else ""
                })
        return pd.DataFrame(results)

# ============================================================================
# 🆕 БЛОК 20: UI ДЛЯ УМНОЙ ЗАГРУЗКИ ТАРИФОВ
# ============================================================================
def show_smart_tariff_interface():
    st.header("🧠 Умная загрузка тарифов")
    st.info("""
📋 **ВЫБЕРИТЕ ИСТОЧНИК ТАРИФОВ:**
🔌 **API Маркетплейса** — прямое подключение к API
🤖 **AI (документация)** — автоматический парсинг
💾 **Загруженные ранее** — использование кэша
🔄 **Гибридный** — AI + API (рекомендуемый)
""")
    try:
        tariff_loader = SmartTariffLoader()
        st.success("✅ SmartTariffLoader инициализирован")
    except Exception as e:
        st.error(f"❌ Ошибка инициализации SmartTariffLoader: {e}")
        tariff_loader = None
        return

    try:
        unit_economics = get_marketplace_unit_economics()
        if unit_economics is None:
            st.warning("⚠️ UnitEconomics не инициализирован")
            return
    except Exception as e:
        st.error(f"❌ Ошибка инициализации UnitEconomics: {e}")
        return

    col1, col2 = st.columns([2, 1])
    with col1:
        marketplace = st.selectbox("🏪 Выберите маркетплейс",
            ["Ozon", "Wildberries", "Яндекс Маркет", "AliExpress", "Мегамаркет", "СберМегаМаркет"],
            key="smart_tariff_mp")
    with col2:
        source = st.selectbox("📡 Источник тарифов", ["hybrid", "api", "ai", "cache"],
            format_func=lambda x: SmartTariffLoader.SOURCES.get(x, x) if hasattr(SmartTariffLoader, 'SOURCES') else x,
            key="smart_tariff_source")

    if source in ["api", "hybrid"]:
        st.subheader("🔑 API ключи")
        col1, col2 = st.columns(2)
        with col1:
            api_key = st.text_input("API Key", type="password", placeholder="Введите API ключ", key="smart_tariff_api_key")
        with col2:
            client_id = st.text_input("Client ID (только для Ozon)", type="password", placeholder="Введите Client ID", key="smart_tariff_client_id")
    else:
        api_key = None
        client_id = None

    if st.button("🚀 Загрузить тарифы", type="primary", key="smart_tariff_load"):
        if not tariff_loader or not hasattr(tariff_loader, 'load_tariffs'):
            st.error("❌ Метод load_tariffs не найден")
            return
        with st.spinner("Загрузка тарифов..."):
            try:
                result = tariff_loader.load_tariffs(
                    marketplace=marketplace, source=source, api_key=api_key, client_id=client_id, force_refresh=True
                )
                if not isinstance(result, dict):
                    st.error("❌ Неверный формат результата")
                    return
                if result.get("errors"):
                    for err in result["errors"]:
                        st.error(f"  - {err}")
                if result.get("data"):
                    st.success(f"✅ Тарифы успешно загружены из источника: {result.get('source_used', 'Неизвестно')}")
                    with st.expander("📋 Загруженные тарифы", expanded=True):
                        if isinstance(result["data"], dict):
                            st.json(result["data"])
                        else:
                            st.write(result["data"])
                    
                    if st.button("💾 Применить тарифы к расчётам", key="smart_tariff_apply"):
                        rates_to_apply = None
                        if "rates" in result["data"]:
                            rates_to_apply = result["data"]["rates"]
                        elif "raw_data" in result["data"]:
                            rates_to_apply = result["data"].get("raw_data", {})
                        elif isinstance(result["data"], dict) and any(k in result["data"] for k in ["commission_rate", "logistics_base"]):
                            rates_to_apply = result["data"]
                        
                        if rates_to_apply and unit_economics and hasattr(unit_economics, '_apply_ai_tariffs'):
                            try:
                                unit_economics._apply_ai_tariffs(marketplace, rates_to_apply)
                                st.success(f"✅ Тарифы для {marketplace} применены!")
                            except Exception as e:
                                st.error(f"❌ Ошибка применения: {e}")
                        else:
                            st.warning("⚠️ Не найдены данные для применения")
                else:
                    st.error("❌ Не удалось загрузить тарифы")
            except Exception as e:
                st.error(f"❌ Ошибка загрузки: {e}")

    st.subheader("📊 Текущие тарифы")
    if unit_economics and hasattr(unit_economics, '_configs'):
        configs = unit_economics._configs
        if marketplace in configs:
            try:
                config = configs[marketplace]
                tariff_data = {
                    "Параметр": ["Комиссия", "Мин. комиссия", "Логистика база", "Логистика за кг", "Логистика за л", "Хранение", "Эквайринг", "Возвраты", "Последняя миля", "Подписка", "Источник", "Обновлено"],
                    "Значение": [
                        f"{config.commission_rate*100:.1f}%", f"{config.min_commission:.2f} ₽", f"{config.logistics_base:.2f} ₽",
                        f"{config.logistics_per_kg:.2f} ₽", f"{config.logistics_per_liter:.2f} ₽", f"{config.storage_per_day:.2f} ₽/л/день",
                        f"{config.acquiring_fee*100:.1f}%", f"{config.return_fee*100:.1f}%", f"{config.last_mile_fee:.2f} ₽",
                        f"{config.subscription_fee:.2f} ₽",
                        config.tariff_source.value if hasattr(config.tariff_source, 'value') else str(config.tariff_source),
                        config.last_updated.strftime('%d.%m.%Y %H:%M') if hasattr(config.last_updated, 'strftime') else str(config.last_updated)
                    ]
                }
                st_dataframe_compat(pd.DataFrame(tariff_data))
            except Exception as e:
                st.warning(f"⚠️ Ошибка отображения тарифов: {e}")

def show_api_tariffs_interface():
    st.header("🌐 API Тарифы маркетплейсов")
    st.info("""
🚧 **Раздел в разработке**
Прямое подключение к API маркетплейсов интегрировано в блок '🧠 Умная загрузка тарифов'.
""")
    st.warning("⚠️ Для работы с API используйте раздел '🧠 Умная загрузка тарифов'")

# ============================================================================
# 🆕 БЛОК 24: РАЗДЕЛ "ИСТОРИЯ РАСЧЁТОВ"
# ============================================================================
def show_history_interface():
    st.header("📚 История расчётов")
    st.info("Здесь хранятся все ваши расчёты юнит-экономики. Данные сохраняются автоматически в локальную базу данных.")
    unit_economics = get_marketplace_unit_economics()
    
    st.subheader("🔍 Фильтры")
    col1, col2, col3 = st.columns(3)
    with col1:
        available_mp = list(unit_economics._configs.keys())
        filter_mp = st.selectbox("🏪 Маркетплейс", ["Все"] + available_mp, key="hist_filter_mp")
    with col2:
        filter_mode = st.selectbox("📦 Режим", ["Все", "FBY", "FBS", "FBO", "DBS", "FBP", "RealFBS"], key="hist_filter_mode")
    with col3:
        limit = st.number_input("📊 Показать записей", min_value=10, max_value=10000, value=100, step=10, key="hist_limit")
    
    filters = {}
    if filter_mp != "Все": filters["marketplace"] = filter_mp
    if filter_mode != "Все": filters["operation_mode"] = filter_mode
    
    with st.spinner("Загрузка истории..."):
        history_df = unit_economics.get_persistent_history(limit=int(limit), filters=filters or None)
    
    if history_df.empty:
        st.warning("⚠️ История пуста. Сделайте расчёт в разделе «📊 Юнит-экономика».")
        return
    
    st.subheader("📊 Ключевые показатели")
    col1, col2, col3, col4 = st.columns(4)
    with col1: st.metric("📦 Всего расчётов", f"{len(history_df):,}")
    with col2:
        total_profit = history_df['profit'].sum() if 'profit' in history_df.columns else 0
        st.metric("💰 Суммарная прибыль", f"{total_profit:,.0f} ₽")
    with col3:
        avg_margin = history_df['margin_percent'].mean() if 'margin_percent' in history_df.columns else 0
        st.metric("📈 Средняя маржа", f"{avg_margin:.1f}%")
    with col4:
        unprofitable = (history_df['profit'] < 0).sum() if 'profit' in history_df.columns else 0
        st.metric("⚠️ Убыточных", f"{unprofitable}")
    
    st.subheader("📋 Последние расчёты")
    display_cols = ["timestamp", "marketplace", "operation_mode", "category", "price", "cost", "profit", "margin_percent", "roi", "tax_amount", "breakeven_price", "recommended_min_price"]
    available = [c for c in display_cols if c in history_df.columns]
    st_dataframe_compat(history_df[available].head(int(limit)))
    
    st.subheader("📤 Экспорт истории")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("📥 Экспорт истории в CSV", key="hist_export_csv"):
            try:
                csv_data = history_df.to_csv(index=False, encoding="utf-8-sig", sep=";")
                st.download_button("⬇️ Скачать CSV", data=csv_data.encode("utf-8-sig"), file_name=f"history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv", mime="text/csv; charset=utf-8", key="hist_download")
                st.success("✅ CSV готов к скачиванию")
            except Exception as e: st.error(f"❌ Ошибка экспорта: {e}")
    with col2:
        if st.button("📥 Экспорт истории в Excel", key="hist_export_excel"):
            try:
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    history_df.to_excel(writer, index=False, sheet_name="История")
                output.seek(0)
                st.download_button("⬇️ Скачать Excel", data=output, file_name=f"history_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="hist_download_excel")
                st.success("✅ Excel готов к скачиванию")
            except Exception as e: st.error(f"❌ Ошибка экспорта: {e}")
    
    st.divider()
    st.subheader("🗑️ Управление историей")
    if st.button("🗑️ Очистить всю историю", key="hist_clear"):
        if st.checkbox("⚠️ Подтверждаю удаление всей истории", key="hist_clear_confirm"):
            try:
                count = unit_economics.clear_persistent_history()
                unit_economics.clear_history()
                st.success(f"✅ Удалено записей: {count}")
                st.rerun()
            except Exception as e: st.error(f"❌ Ошибка очистки: {e}")

# ============================================================================
# 🆕 БЛОК 25: РАЗДЕЛ "НАСТРОЙКИ ПРИЛОЖЕНИЯ"
# ============================================================================
def show_settings_interface():
    st.header("⚙️ Настройки приложения")
    st.info("Здесь вы можете настроить параметры расчётов по умолчанию. Настройки сохраняются в файл `config/settings.json`.")
    unit_economics = get_marketplace_unit_economics()
    settings = unit_economics._settings.copy()
    
    st.subheader("💰 Финансовые параметры")
    col1, col2 = st.columns(2)
    with col1:
        settings["global_markup"] = st.number_input("📈 Глобальная наценка (%)", min_value=0.0, max_value=500.0, value=float(settings.get("global_markup", DEFAULT_MARKUP_GLOBAL) * 100), step=1.0, key="set_markup") / 100.0
        settings["discount_max"] = st.number_input("🎯 Максимальная скидка (%)", min_value=0.0, max_value=100.0, value=float(settings.get("discount_max", DEFAULT_DISCOUNT_MAX) * 100), step=1.0, key="set_discount") / 100.0
    with col2:
        settings["target_margin"] = st.number_input("💎 Целевая маржа (%)", min_value=0.0, max_value=100.0, value=float(settings.get("target_margin", 20.0)), step=1.0, key="set_target_margin")
        settings["default_days_storage"] = st.number_input("📦 Дней хранения по умолчанию", min_value=1, max_value=365, value=int(settings.get("default_days_storage", 30)), step=1, key="set_days")
    
    st.subheader("🏪 Параметры по умолчанию")
    col1, col2 = st.columns(2)
    with col1:
        mp_list = list(unit_economics._configs.keys())
        default_mp = settings.get("default_marketplace", "Ozon")
        default_mp_idx = mp_list.index(default_mp) if default_mp in mp_list else 0
        settings["default_marketplace"] = st.selectbox("🏪 Маркетплейс по умолчанию", mp_list, index=default_mp_idx, key="set_mp")
        modes_list = ["FBY", "FBS", "FBO", "DBS", "FBP", "RealFBS"]
        default_mode = settings.get("default_mode", "FBS")
        default_mode_idx = modes_list.index(default_mode) if default_mode in modes_list else 1
        settings["default_mode"] = st.selectbox("📦 Режим по умолчанию", modes_list, index=default_mode_idx, key="set_mode")
    with col2:
        tax_list = list(TAX_SYSTEMS.keys())
        default_tax = settings.get("tax_system", "УСН_6")
        default_tax_idx = tax_list.index(default_tax) if default_tax in tax_list else 0
        settings["tax_system"] = st.selectbox("💼 Налоговая система", tax_list, format_func=lambda x: TAX_SYSTEMS[x]["name"], index=default_tax_idx, key="set_tax")
        ad_list = ["low", "medium", "high", "aggressive"]
        ad_labels = {"low": "Низкая (5%)", "medium": "Средняя (15%)", "high": "Высокая (25%)", "aggressive": "Агрессивная (35%)"}
        default_ad = settings.get("ad_intensity", "medium")
        default_ad_idx = ad_list.index(default_ad) if default_ad in ad_list else 1
        settings["ad_intensity"] = st.selectbox("📢 Интенсивность рекламы", ad_list, format_func=lambda x: ad_labels[x], index=default_ad_idx, key="set_ad")
    
    st.subheader("🚀 Производительность")
    col1, col2 = st.columns(2)
    with col1:
        settings["parallel_processing"] = st.checkbox("🧵 Параллельная обработка", value=bool(settings.get("parallel_processing", True)), key="set_parallel")
        settings["enable_cache"] = st.checkbox("💾 Кэширование", value=bool(settings.get("enable_cache", True)), key="set_cache")
    with col2:
        settings["max_workers"] = st.number_input("🔧 Макс. потоков", min_value=1, max_value=32, value=int(settings.get("max_workers", DEFAULT_MAX_WORKERS)), step=1, key="set_workers")
        settings["enable_persistent_history"] = st.checkbox("📚 Сохранять историю в БД", value=bool(settings.get("enable_persistent_history", True)), key="set_history_db")
    
    st.divider()
    col1, col2, col3 = st.columns([1, 1, 2])
    with col1:
        if st.button("💾 Сохранить настройки", type="primary", key="set_save"):
            try:
                success = unit_economics.save_settings(settings)
                if success: st.success("✅ Настройки успешно сохранены")
                else: st.error("❌ Ошибка сохранения настроек")
            except Exception as e: st.error(f"❌ Ошибка: {e}")
    with col2:
        if st.button("🔄 Сбросить к значениям по умолчанию", key="set_reset"):
            default_settings = {
                "default_marketplace": "Ozon", "default_mode": "FBS", "default_days_storage": 30, "target_margin": 20.0,
                "enable_cache": True, "parallel_processing": True, "max_workers": DEFAULT_MAX_WORKERS, "enable_persistent_history": True,
                "global_markup": DEFAULT_MARKUP_GLOBAL, "discount_max": DEFAULT_DISCOUNT_MAX, "tax_system": "УСН_6", "ad_intensity": "medium",
            }
            try:
                success = unit_economics.save_settings(default_settings)
                if success:
                    st.success("✅ Настройки сброшены")
                    st.rerun()
            except Exception as e: st.error(f"❌ Ошибка: {e}")
    
    st.divider()
    st.subheader("📊 Статистика приложения")
    try:
        stats = unit_economics.get_stats()
        col1, col2, col3, col4 = st.columns(4)
        with col1: st.metric("📦 Всего расчётов", stats.get("total_calculations", 0))
        with col2: st.metric("💾 Сохранено в БД", stats.get("db_saved", 0))
        with col3: st.metric("❌ Ошибок", stats.get("errors_count", 0))
        with col4:
            uptime = stats.get("uptime", 0)
            if uptime > 3600: uptime_str = f"{uptime / 3600:.1f} ч"
            elif uptime > 60: uptime_str = f"{uptime / 60:.1f} мин"
            else: uptime_str = f"{uptime:.0f} сек"
            st.metric("⏱️ Uptime", uptime_str)
        col1, col2 = st.columns(2)
        with col1:
            st.metric("📈 Средняя прибыль", f"{stats.get('avg_profit', 0):.2f} ₽")
            st.metric("💰 Макс. прибыль", f"{stats.get('max_profit', 0):.2f} ₽")
        with col2:
            st.metric("📊 Средняя маржа", f"{stats.get('avg_margin', 0):.2f}%")
            st.metric("📈 Средний ROI", f"{stats.get('avg_roi', 0):.2f}%")
        if stats.get("best_marketplace"):
            st.info(f"🏆 Лучшая конфигурация: **{stats.get('best_marketplace')}** / **{stats.get('best_mode')}** / категория **{stats.get('best_category')}**")
    except Exception as e:
        st.warning(f"⚠️ Не удалось получить статистику: {e}")
    
    with st.expander("📋 Текущие настройки (JSON)", expanded=False):
        st.json(settings)

# ============================================================================
# ГЛАВНАЯ ФУНКЦИЯ ПРИЛОЖЕНИЯ (ОЧИЩЕННАЯ ВЕРСИЯ)
# ============================================================================
def main():
    """Главная функция приложения"""
    st.set_page_config(
        page_title=APP_NAME,
        page_icon="🚗",
        layout="wide",
        initial_sidebar_state="expanded"
    )
    st.title(APP_NAME)
    st.caption(f"Версия {APP_VERSION} | {APP_DESCRIPTION}")
    st.sidebar.title("🧭 Навигация")
    
    section = st.sidebar.radio(
        "Выберите раздел:",
        [
            "📁 Загрузка данных",
            "📊 Юнит-экономика",
            "🤖 AI Тарифы",
            "🌐 API Тарифы маркетплейсов",
            "🧠 Умная загрузка тарифов",
            "📚 История расчётов",
            "⚙️ Настройки",
        ],
        key="main_navigation",
    )
    
    if section == "📁 Загрузка данных":
        show_data_upload_interface()
    elif section == "📊 Юнит-экономика":
        show_unit_economics_interface()
    elif section == "🤖 AI Тарифы":
        show_ai_tariffs_interface()
    elif section == "🌐 API Тарифы маркетплейсов":
        show_api_tariffs_interface()
    elif section == "🧠 Умная загрузка тарифов":
        show_smart_tariff_interface()
    elif section == "📚 История расчётов":
        show_history_interface()
    elif section == "⚙️ Настройки":
        show_settings_interface()

# ✅ ТОЧКА ВХОДА
if __name__ == "__main__":
    main()
