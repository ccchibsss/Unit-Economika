# -*- coding: utf-8 -*-
"""
🚗 Юнит-экономика FBS — Яндекс Маркет | Монолит Streamlit v4.0
Запуск:
    pip install streamlit pandas numpy plotly requests xlsxwriter openpyxl
    streamlit run streamlit_app.py
Функционал 1-в-1 с веб-версией:
  • Бренд + Артикул + Категория + Длина/Ширина/Высота + Цена (+ себестоимость, вес)
  • Себестоимость необязательна → % от цены, вес по плотности, справочник категорий
  • Базовый тариф, спецтарифы (шины/АКБ/двигатели/КПП), доп.расходы, плотность
  • Потоковое чтение CSV, расчёт чанками, прогресс-бары, до 300 000 SKU
  • Дашборд с аналитикой по категориям И брендам, виртуализация не нужна — pandas
  • Экспорт: Excel с живыми формулами (≤50k), быстрый Excel со значениями (≤300k),
    только сводки, CSV, шаблон
"""
import io
import json
import logging
import math
import time
import re
from dataclasses import asdict, dataclass, field
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Dict, List, Tuple
import numpy as np
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import requests
import streamlit as st
# ─────────────────────────────────────────────────────────────────────────────
# КОНСТАНТЫ
# ─────────────────────────────────────────────────────────────────────────────
APP_VERSION = "4.0.0"
APP_NAME = "🚗 Юнит-экономика FBS — Яндекс Маркет"
BIG_DATA_THRESHOLD = 20_000
FORMULA_ROW_LIMIT = 50_000
VALUES_ROW_LIMIT = 300_000
CHUNK_SIZE = 25_000
DEFAULT_COST_FALLBACK = 0.65
DEFAULT_DENSITY = 0.30
DEFAULT_BASE = dict(commission_rate=0.14, min_commission=45.0, logistics_base=45.0,
                    logistics_per_kg=14.0, storage_per_day_per_liter=0.25,
                    acquiring_fee=0.02, return_fee=0.02)
DEFAULT_SPECIAL_COSTS = dict(packaging=45.0, chestny_znak=1.5, labeling=3.0,
                             warranty_reserve=0.02, hazard_surcharge=0.01, fragile_surcharge=0.005)
SPECIAL_TARIFFS = {
    "шины":        dict(commission_rate=0.12, logistics_base=90.0,  storage_per_day_per_liter=0.50, label="Шины"),
    "аккумулятор": dict(commission_rate=0.13, logistics_base=75.0,  storage_per_day_per_liter=0.40, label="Аккумуляторы"),
    "двигател":    dict(commission_rate=0.11, logistics_base=120.0, storage_per_day_per_liter=0.60, label="Двигатели"),
    "кпп":         dict(commission_rate=0.11, logistics_base=110.0, storage_per_day_per_liter=0.60, label="КПП"),
}
CATEGORIES_DB = {
    "фильтры":      (1.5,  0.5, False, False),
    "масла":        (5.0,  4.0, True,  False),
    "колодки":      (0.8,  1.2, False, False),
    "диски":        (3.0,  4.0, False, True),
    "амортизаторы": (4.0,  3.5, False, True),
    "аккумуляторы": (12.0, 15.0, True, True),
    "шины":         (25.0, 10.0, False, False),
    "фары":         (6.0,  2.5, False, True),
    "двигател":     (50.0, 80.0, True, True),
    "кпп":          (40.0, 50.0, True, True),
}
FALLBACK_CAT = (2.0, 1.0, False, False)
DEMO_BRANDS = ["Bosch","Mann-Filter","Sachs","Brembo","Mahle","Denso","Valeo","TRW","NGK",
               "Febi Bilstein","Lemförder","Hella","Continental","Michelin","Varta"]
DEMO_CATS = [
    ("Фильтры",      280,  900,  22,14,14, 0.6, 25),
    ("Масла",        700,  4200, 28,16,28, 4.4, 20),
    ("Колодки",      1500, 3800, 18,12, 8, 1.4, 30),
    ("Диски",        2500, 9500, 62,62,22, 9.2, 45),
    ("Амортизаторы", 3200, 7800, 66,16,16, 3.6, 40),
    ("Аккумуляторы", 5200,12500, 35,26,26,17.0, 35),
    ("Шины",         4200,14500, 70,70,26,10.6, 15),
    ("Фары",         2800,18500, 52,26,26, 2.6, 50),
    ("Двигатели",   65000,180000,100,62,72,95.0, 90),
    ("КПП",         48000,120000,82,58,56,50.0, 90),
]
SYNONYMS = {
    "Артикул": ["артикул","sku","код","offer id","offerid","shop-sku","артикул товара"],
    "Бренд": ["бренд","brand","производитель","марка","торговая марка","vendor"],
    "Категория": ["категория","category","группа","тип товара","раздел"],
    "Цена": ["цена","price","розничная цена","цена продажи"],
    "Себестоимость": ["себестоимость","cost","закупка","закупочная цена","закуп"],
    "Вес_кг": ["вес_кг","вес","weight","масса"],
    "Длина": ["длина","length"],
    "Ширина": ["ширина","width"],
    "Высота": ["высота","height"],
    "Объем_л": ["объем_л","объем","объём","volume"],
    "Оборачиваемость_дней": ["оборачиваемость","turnover","срок хранения"],
    "Опасный": ["опасный","hazardous"],
    "Хрупкий": ["хрупкий","fragile"],
}
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("FBS")
# ─────────────────────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────────────────────
CSS = r"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700;800&family=Manrope:wght@700;800&display=swap');
html, body, [class*="css"] { font-family: Inter, sans-serif; }
.block-container { padding-top: 1rem; max-width: 1320px; }
.hero {
  background: linear-gradient(135deg,#0f172a 0%,#1e1b4b 40%,#312e81 72%,#4338ca 100%);
  border-radius: 20px; padding: 22px 24px; color: white; position: relative; overflow:hidden;
}
.hero:before { content:""; position:absolute; inset:0;
  background: radial-gradient(600px 200px at 88% 10%, rgba(251,146,60,.22), transparent 70%); }
.hero h1 { font-family: Manrope, sans-serif; font-size: 1.65rem; font-weight: 800; margin:0; position:relative; }
.hero p { color:#c7d2fe; font-size:.9rem; margin:4px 0 0; position:relative; }
.badge { display:inline-flex; align-items:center; gap:6px; background: rgba(255,255,255,.12);
  border:1px solid rgba(255,255,255,.18); padding:4px 10px; border-radius:999px; font-size:.74rem; font-weight:600; }
.card { background:white; border:1px solid #e2e8f0; border-radius:16px; padding:16px; }
.metric-card { background:white; border:1px solid #e2e8f0; border-radius:16px; padding:14px 16px; position:relative; overflow:hidden; }
.metric-card:before { content:""; position:absolute; left:0; right:0; top:0; height:3px; }
.mc-indigo:before{ background: linear-gradient(90deg,#6366f1,#8b5cf6);}
.mc-emerald:before{ background: linear-gradient(90deg,#10b981,#06b6d4);}
.mc-rose:before{ background: linear-gradient(90deg,#f43f5e,#ec4899);}
.mc-amber:before{ background: linear-gradient(90deg,#f59e0b,#f97316);}
.mc-violet:before{ background: linear-gradient(90deg,#8b5cf6,#a855f7);}
.stepper { display:flex; gap:8px; flex-wrap:wrap; }
.step { flex:1; min-width:150px; border-radius:14px; padding:10px 12px; border:1px solid #e2e8f0; background:white; cursor:pointer; }
.step.active { border-color:#6366f1; background: linear-gradient(135deg,#eef2ff,#f5f3ff); }
.step.done { border-color:#a7f3d0; background:#ecfdf5; }
.small { font-size:.78rem; color:#64748b; }
</style>
"""
# ─────────────────────────────────────────────────────────────────────────────
# УТИЛИТЫ
# ─────────────────────────────────────────────────────────────────────────────
def money(v): return f"{v:,.0f} ₽".replace(","," ")
def money2(v): return f"{v:,.2f} ₽".replace(","," ")
def pct(v, d=1): return f"{v*100:.{d}f}%".replace(".",",")
def fmt_size(b):
    return f"{b/1024/1024:.1f} МБ" if b>1024*1024 else f"{b/1024:.0f} КБ"
def to_num(v):
    if isinstance(v,(int,float)) and np.isfinite(v): return float(v)
    if v is None or v=="": return 0.0
    s=str(v).strip().replace("\u00a0","").replace(" ","").replace("₽","").replace("%","")
    if not s or s=="-": return 0.0
    if "," in s and "." not in s: s=s.replace(",",".")
    else: s=s.replace(",","")
    try: return float(s)
    except: return 0.0
def nullable_num(v):
    if v is None or v=="": return None
    n=to_num(v)
    return None if n==0 else n
TRUE_SET={"да","yes","true","1","есть","+"}
FALSE_SET={"нет","no","false","0","-"}
def to_bool(v):
    if v is None or v=="": return None
    if isinstance(v,bool): return v
    s=str(v).strip().lower()
    if s in TRUE_SET: return True
    if s in FALSE_SET: return False
    return None
def get_cat_defaults(cat_lower:str):
    for k,(vol,w,ha,fr) in CATEGORIES_DB.items():
        if k in cat_lower: return vol,w,ha,fr
    return FALLBACK_CAT
def mulberry32(seed):
    def rnd():
        nonlocal seed
        seed = (seed + 0x6d2b79f5) & 0xffffffff
        t = (seed ^ (seed >> 15)) * (1 | seed) & 0xffffffff
        t = (t + ((t ^ (t >> 7)) * (61 | t) & 0xffffffff)) & 0xffffffff
        t ^= t >> 14
        return (t & 0xffffffff) / 4294967296
    return rnd
def build_demo_rows(count=24, with_cost=True):
    rnd=mulberry32(20260408+count)
    rows=[]
    for i in range(count):
        name,p0,p1,d0,d1,d2,w0,turn = DEMO_CATS[i % len(DEMO_CATS)]
        brand = DEMO_BRANDS[(i*7+i%3) % len(DEMO_BRANDS)]
        price = round((p0 + rnd()*(p1-p0))/10)*10
        j = 0.85 + rnd()*0.35
        rows.append(dict(
            Артикул=f"{brand[:3].upper()}-{name[:3].upper()}-{i+1:06d}",
            Бренд=brand, Категория=name,
            Цена=price,
            Себестоимость=round(price*(0.55+rnd()*0.22)) if with_cost else None,
            Вес_кг=round(w0*(0.8+rnd()*0.45),2),
            Длина=round(d0*j), Ширина=round(d1*j), Высота=round(d2*j),
            Объем_л=None, Оборачиваемость_дней=round(turn*(0.7+rnd()*0.7)),
            Опасный=None, Хрупкий=None
        ))
    return pd.DataFrame(rows)
# простая модель тарифа в session_state
def default_tariff():
    return dict(
        base=dict(DEFAULT_BASE),
        special_enabled=True,
        special_costs=dict(DEFAULT_SPECIAL_COSTS),
        cost_fallback=DEFAULT_COST_FALLBACK,
        density=DEFAULT_DENSITY,
        category_rates={},
        use_category_rates=True,
    )
# ─────────────────────────────────────────────────────────────────────────────
# ПАРСИНГ ФАЙЛА (синонимы → канонические колонки)
# ─────────────────────────────────────────────────────────────────────────────
def resolve_columns(headers):
    out={}
    low=[h.strip().lower() for h in headers]
    for canon,syns in SYNONYMS.items():
        found=None
        for i,h in enumerate(headers):
            if low[i] in syns: found=h; break
        if not found:
            for i,h in enumerate(headers):
                if any(s in low[i] for s in syns): found=h; break
        out[canon]=found
    return out
def dim_factor(header):
    if not header: return 1.0
    h=header.lower()
    if "мм" in h or "mm" in h: return 0.1
    if "метр" in h: return 100.0
    return 1.0
def parse_dataframe(df_raw: pd.DataFrame, file_name: str, file_size:int):
    t0=time.time()
    headers=list(df_raw.columns.astype(str))
    cmap=resolve_columns(headers)
    # проверка обязательных
    missing=[k for k in ["Артикул","Категория","Цена"] if not cmap[k]]
    has_cost= bool(cmap["Себестоимость"])
    has_dims= bool(cmap["Длина"] and cmap["Ширина"] and cmap["Высота"])
    # быстрый маппинг
    rename={v:k for k,v in cmap.items() if v}
    df=df_raw.rename(columns=rename)
    # оставляем только канон
    keep=[c for c in ["Артикул","Бренд","Категория","Цена","Себестоимость","Вес_кг","Длина","Ширина","Высота","Объем_л","Оборачиваемость_дней","Опасный","Хрупкий"] if c in df.columns]
    df=df[keep].copy()
    if "Бренд" not in df.columns: df["Бренд"]=""
    if "Артикул" not in df.columns: df["Артикул"]=""
    if "Категория" not in df.columns: df["Категория"]=""
    # числовые
    for c in ["Цена","Себестоимость","Вес_кг","Длина","Ширина","Высота","Объем_л","Оборачиваемость_дней"]:
        if c in df.columns: df[c]=df[c].apply(lambda x: to_num(x) if pd.notna(x) and str(x).strip()!="" else np.nan)
    # габариты в см
    for c in ["Длина","Ширина","Высота"]:
        if c in df.columns and cmap[c]:
            f=dim_factor(cmap[c]); df[c]=df[c]*f
    # bool
    for c in ["Опасный","Хрупкий"]:
        if c in df.columns: df[c]=df[c].apply(to_bool)
    # убираем полностью пустые
    mask = (df["Артикул"].astype(str).str.strip()!="") | (df["Категория"].astype(str).str.strip()!="") | (df["Цена"].fillna(0)!=0)
    skipped = int((~mask).sum())
    df=df[mask].reset_index(drop=True)
    if df.empty: raise ValueError("После очистки не осталось строк с данными")
    # пустые артикулы
    empty_art = df["Артикул"].astype(str).str.strip()==""
    df.loc[empty_art,"Артикул"]=[f"SKU-{i+1}" for i in df[empty_art].index]
    # себестоимость: NaN оставим — дальше подставим fallback
    # NaN → None для совместимости
    df=df.replace({np.nan: None})
    parse_ms=(time.time()-t0)*1000
    return df, cmap, missing, has_cost, has_dims, skipped, parse_ms
# ─────────────────────────────────────────────────────────────────────────────
# РАСЧЁТ (векторизованный, один проход)
# ─────────────────────────────────────────────────────────────────────────────
def effective_for_category(cat_lower:str, tariff:dict):
    base=tariff["base"]
    eff=dict(base)
    eff["special_applied"]=False; eff["reason"]=""
    if tariff["special_enabled"]:
        for key,rule in SPECIAL_TARIFFS.items():
            if key in cat_lower:
                eff["commission_rate"]=rule["commission_rate"]
                eff["logistics_base"]=rule["logistics_base"]
                eff["storage_per_day_per_liter"]=rule["storage_per_day_per_liter"]
                eff["special_applied"]=True
                if "шины" in cat_lower: eff["reason"]="Крупногабаритный"
                elif "аккумулятор" in cat_lower: eff["reason"]="Опасный груз"
                else: eff["reason"]="Крупногабаритный/тяжёлый"
                break
    if tariff["use_category_rates"] and tariff["category_rates"]:
        for k,v in tariff["category_rates"].items():
            if k.lower() in cat_lower:
                eff["commission_rate"]=v; break
    return eff
def calculate_df(df: pd.DataFrame, tariff:dict, progress_cb=None) -> pd.DataFrame:
    n=len(df)
    # кэш тарифов по категориям
    cats=df["Категория"].fillna("").astype(str).str.lower().unique()
    cache={c: effective_for_category(c,tariff) for c in cats}
    # массивы для векторизации
    cat_low=df["Категория"].fillna("").astype(str).str.lower()
    commission_rate=np.array([cache[c]["commission_rate"] for c in cat_low])
    min_comm=np.array([cache[c]["min_commission"] for c in cat_low])
    log_base=np.array([cache[c]["logistics_base"] for c in cat_low])
    log_per_kg=np.array([cache[c]["logistics_per_kg"] for c in cat_low])
    stor_rate=np.array([cache[c]["storage_per_day_per_liter"] for c in cat_low])
    acq=np.array([cache[c]["acquiring_fee"] for c in cat_low])
    price=df["Цена"].fillna(0).astype(float).values
    cost_raw=df["Себестоимость"].values  # может быть None
    # себестоимость: где None/0 → fallback
    cost=np.array([ (c if c not in (None,0) else p*tariff["cost_fallback"]) for c,p in zip(cost_raw,price)], dtype=float)
    cost_estimated=np.array([ c in (None,0) for c in cost_raw])
    # габариты
    L=df["Длина"].fillna(0).astype(float).values if "Длина" in df.columns else np.zeros(n)
    W=df["Ширина"].fillna(0).astype(float).values if "Ширина" in df.columns else np.zeros(n)
    H=df["Высота"].fillna(0).astype(float).values if "Высота" in df.columns else np.zeros(n)
    has_dims=(L>0)&(W>0)&(H>0)
    volume_raw=df["Объем_л"].fillna(0).astype(float).values if "Объем_л" in df.columns else np.zeros(n)
    # объём
    volume=np.where(volume_raw>0, volume_raw, np.where(has_dims, L*W*H/1000, 0))
    # дефолты по категориям где объём 0
    for i in range(n):
        if volume[i]==0:
            vol,_,_,_=get_cat_defaults(cat_low[i])
            volume[i]=vol
    weight_raw=df["Вес_кг"].fillna(0).astype(float).values if "Вес_кг" in df.columns else np.zeros(n)
    # перевод граммов
    weight_raw=np.where(weight_raw>100, weight_raw/1000, weight_raw)
    weight=np.where(weight_raw>0, weight_raw, 0)
    # оценка по объёму
    weight=np.where((weight==0)&has_dims, np.maximum(0.1, volume*tariff["density"]), weight)
    for i in range(n):
        if weight[i]==0:
            _,w,_,_=get_cat_defaults(cat_low[i]); weight[i]=w
    volumetric=np.where(has_dims, L*W*H/5000, 0)
    billable=np.maximum.reduce([weight, volumetric, np.full(n,0.1)])
    # флаги опасности/хрупкости
    haz_raw=df["Опасный"].values if "Опасный" in df.columns else [None]*n
    fra_raw=df["Хрупкий"].values if "Хрупкий" in df.columns else [None]*n
    haz=np.zeros(n, dtype=bool); fra=np.zeros(n, dtype=bool)
    for i in range(n):
        _,_,dh,fr=get_cat_defaults(cat_low[i])
        haz[i]= haz_raw[i] if haz_raw[i] is not None else dh
        fra[i]= fra_raw[i] if fra_raw[i] is not None else fr
    turnover=df["Оборачиваемость_дней"].fillna(30).astype(float).values if "Оборачиваемость_дней" in df.columns else np.full(n,30)
    turnover=np.where(turnover<=0,30,turnover)
    sc=tariff["special_costs"]
    spec_costs = sc["packaging"]+sc["chestny_znak"]+sc["labeling"] + price*sc["warranty_reserve"] + np.where(haz, price*sc["hazard_surcharge"],0) + np.where(fra, price*sc["fragile_surcharge"],0)
    commission=np.maximum(price*commission_rate, min_comm)
    logistics=log_base + billable*log_per_kg
    storage=volume*stor_rate*turnover
    acquiring=price*acq
    total=cost+commission+logistics+storage+acquiring+spec_costs
    profit=price-total
    margin=np.where(price>0, profit/price, 0)
    out=df.copy()
    out["Бренд"]=out["Бренд"].fillna("").replace("", "Без бренда")
    out["Категория"]=out["Категория"].fillna("").replace("", "Без категории")
    out["Цена"]=price
    out["Себестоимость"]=cost
    out["Себестоимость_оценка"]=cost_estimated
    out["Вес_кг"]=weight
    out["Длина"]=L; out["Ширина"]=W; out["Высота"]=H
    out["Объем_л"]=volume
    out["Оплач_вес"]=billable
    out["Оборачиваемость_дней"]=turnover
    out["is_hazardous"]=haz; out["is_fragile"]=fra
    out["Спец_расходы_FBS"]=spec_costs
    out["Комиссия_руб"]=commission
    out["Логистика_руб"]=logistics
    out["Хранение_руб"]=storage
    out["Эквайринг_руб"]=acquiring
    out["Итого_расходы"]=total
    out["Прибыль"]=profit
    out["Маржа_%"]=margin
    out["Спецтариф_применён"]=np.array([cache[c]["special_applied"] for c in cat_low])
    out["Причина_спецтарифа"]=np.array([cache[c]["reason"] for c in cat_low])
    return out
def totals_row(df: pd.DataFrame):
    rev=float(df["Цена"].sum()); exp=float(df["Итого_расходы"].sum()); prof=float(df["Прибыль"].sum())
    return dict(revenue=rev, expenses=exp, profit=prof, avg_margin=prof/rev if rev else 0,
                count=len(df), loss=int((df["Прибыль"]<0).sum()),
                spec=int(df["Спецтариф_применён"].sum()),
                est=int(df["Себестоимость_оценка"].sum()) if "Себестоимость_оценка" in df.columns else 0,
                commission=float(df["Комиссия_руб"].sum()),
                logistics=float(df["Логистика_руб"].sum()),
                storage=float(df["Хранение_руб"].sum()),
                cost=float(df["Себестоимость"].sum()))
def summarize(df: pd.DataFrame, by:str):
    g=df.groupby(by, dropna=False).agg(
        count=("Артикул","count"), revenue=("Цена","sum"),
        expenses=("Итого_расходы","sum"), profit=("Прибыль","sum"),
        spec=("Спецтариф_применён","sum"), loss=("Прибыль", lambda x: (x<0).sum())
    ).reset_index().rename(columns={by:"name"})
    g["avg_margin"]=np.where(g["revenue"]>0, g["profit"]/g["revenue"], 0)
    return g.sort_values("revenue", ascending=False)
# ─────────────────────────────────────────────────────────────────────────────
# ЭКСПОРТ EXCEL
# ─────────────────────────────────────────────────────────────────────────────
def build_excel_formula(df: pd.DataFrame, tariff:dict) -> bytes:
    import xlsxwriter
    out=io.BytesIO()
    wb=xlsxwriter.Workbook(out, {"in_memory": True})
    fmt_h=wb.add_format({"bold":True,"bg_color":"#0F3460","font_color":"#FFFFFF","border":1,"align":"center","valign":"vcenter"})
    fmt_m=wb.add_format({"num_format":"#,##0.00","border":1})
    fmt_p=wb.add_format({"num_format":"0.00%","border":1})
    fmt_red=wb.add_format({"bg_color":"#FFC7CE","border":1})
    fmt_green=wb.add_format({"bg_color":"#C6EFCE","border":1})
    fmt_formula=wb.add_format({"bg_color":"#DCE6F1","border":1})
    # Тариф
    ws=wb.add_worksheet("Тариф")
    ws.write_row(0,0,["Параметр","Значение"], fmt_h)
    rows=[("Комиссия, %",tariff["base"]["commission_rate"],fmt_p),
          ("Мин. комиссия, ₽",tariff["base"]["min_commission"],fmt_m),
          ("Логистика база, ₽",tariff["base"]["logistics_base"],fmt_m),
          ("Логистика за кг, ₽",tariff["base"]["logistics_per_kg"],fmt_m),
          ("Хранение за л/сутки, ₽",tariff["base"]["storage_per_day_per_liter"],fmt_m),
          ("Эквайринг, %",tariff["base"]["acquiring_fee"],fmt_p),
          ("Возвраты, %",tariff["base"]["return_fee"],fmt_p),
          ("Упаковка FBS, ₽",tariff["special_costs"]["packaging"],fmt_m),
          ("Честный знак, ₽",tariff["special_costs"]["chestny_znak"],fmt_m),
          ("Маркировка, ₽",tariff["special_costs"]["labeling"],fmt_m),
          ("Гарант. резерв, %",tariff["special_costs"]["warranty_reserve"],fmt_p),
          ("Надбавка опасный, %",tariff["special_costs"]["hazard_surcharge"],fmt_p),
          ("Надбавка хрупкий, %",tariff["special_costs"]["fragile_surcharge"],fmt_p),
          ("Себест. по умолч., % от цены",tariff["cost_fallback"],fmt_p),
          ("Плотность, кг/л",tariff["density"],wb.add_format({"num_format":"0.00","border":1})),
          ("Спецтарифы", "Применяются" if tariff["special_enabled"] else "Отключены", wb.add_format({"border":1}))]
    for i,(a,b,f) in enumerate(rows, start=1):
        ws.write(i,0,a); ws.write(i,1,b,f)
    ws.set_column(0,0,38); ws.set_column(1,1,16)
    # Входные
    wi=wb.add_worksheet("Входные_Данные")
    cols=["Артикул","Бренд","Категория","Длина","Ширина","Высота","Объем_л","Оплач_вес","Цена","Себестоимость","Спец_расходы_FBS","Оборачиваемость_дней"]
    wi.write_row(0,0,cols, fmt_h)
    for r, (_,row) in enumerate(df.iterrows(), start=1):
        wi.write(r,0,row["Артикул"]); wi.write(r,1,row["Бренд"]); wi.write(r,2,row["Категория"])
        wi.write(r,3,row["Длина"]); wi.write(r,4,row["Ширина"]); wi.write(r,5,row["Высота"])
        wi.write(r,6,round(float(row["Объем_л"]),2)); wi.write(r,7,round(float(row["Оплач_вес"]),2))
        wi.write(r,8,row["Цена"]); wi.write(r,9,round(float(row["Себестоимость"]),2))
        wi.write(r,10,round(float(row["Спец_расходы_FBS"]),2)); wi.write(r,11,row["Оборачиваемость_дней"])
    wi.freeze_panes(1,0); wi.autofilter(0,0,len(df), len(cols)-1)
    wi.set_column(0,1,16); wi.set_column(2,2,20)
    # Расчёт с формулами
    wc=wb.add_worksheet("Расчет_FBS")
    headers=["Артикул","Бренд","Категория","Цена","Себестоимость","Комиссия_руб","Логистика_руб","Хранение_руб","Эквайринг_руб","Спец_расходы_FBS","Итого_расходы","Прибыль","Маржа_%","Спецтариф_применён","Причина_спецтарифа"]
    wc.write_row(0,0,headers, fmt_h)
    IN="'Входные_Данные'"; T="'Тариф'"
    for i in range(len(df)):
        n=i+2
        wc.write_formula(i+1,0, f"={IN}!A{n}", None, df.iloc[i]["Артикул"])
        wc.write_formula(i+1,1, f"={IN}!B{n}", None, df.iloc[i]["Бренд"])
        wc.write_formula(i+1,2, f"={IN}!C{n}", None, df.iloc[i]["Категория"])
        wc.write_formula(i+1,3, f"={IN}!I{n}", None, float(df.iloc[i]["Цена"]))
        wc.write_formula(i+1,4, f"={IN}!J{n}", None, round(float(df.iloc[i]["Себестоимость"]),2))
        wc.write_formula(i+1,5, f"=MAX(D{n}*{T}!$B$2,{T}!$B$3)", None, round(float(df.iloc[i]["Комиссия_руб"]),2))
        wc.write_formula(i+1,6, f"={T}!$B$4+{IN}!H{n}*{T}!$B$5", None, round(float(df.iloc[i]["Логистика_руб"]),2))
        wc.write_formula(i+1,7, f"={IN}!G{n}*{T}!$B$6*{IN}!L{n}", None, round(float(df.iloc[i]["Хранение_руб"]),2))
        wc.write_formula(i+1,8, f"=D{n}*{T}!$B$7", None, round(float(df.iloc[i]["Эквайринг_руб"]),2))
        wc.write_formula(i+1,9, f"={IN}!K{n}", None, round(float(df.iloc[i]["Спец_расходы_FBS"]),2))
        wc.write_formula(i+1,10,f"=E{n}+F{n}+G{n}+H{n}+I{n}+J{n}", None, round(float(df.iloc[i]["Итого_расходы"]),2))
        wc.write_formula(i+1,11,f"=D{n}-K{n}", None, round(float(df.iloc[i]["Прибыль"]),2))
        wc.write_formula(i+1,12,f"=IF(D{n}>0,L{n}/D{n},0)", None, float(df.iloc[i]["Маржа_%"]))
        wc.write_formula(i+1,13,f'=IF(OR(ISNUMBER(SEARCH("шины",C{n})),ISNUMBER(SEARCH("аккумулятор",C{n})),ISNUMBER(SEARCH("двигател",C{n})),ISNUMBER(SEARCH("кпп",C{n}))),TRUE,FALSE)', None, bool(df.iloc[i]["Спецтариф_применён"]))
        wc.write_formula(i+1,14,f'=IF(N{n},IF(ISNUMBER(SEARCH("шины",C{n})),"Крупногабаритный",IF(ISNUMBER(SEARCH("аккумулятор",C{n})),"Опасный груз","Крупногабаритный/тяжёлый")),"")', None, df.iloc[i]["Причина_спецтарифа"])
    wc.set_column(3,11, 12, fmt_m); wc.set_column(12,12, 10, fmt_p)
    wc.freeze_panes(1,0); wc.autofilter(0,0,len(df),14)
    if len(df)>1:
        wc.conditional_format(f"M2:M{len(df)+1}", {"type":"cell","criteria":"<","value":0,"format":fmt_red})
        wc.conditional_format(f"M2:M{len(df)+1}", {"type":"cell","criteria":"between","minimum":0.15,"maximum":10,"format":fmt_green})
        wc.conditional_format(f"L2:L{len(df)+1}", {"type":"cell","criteria":"<","value":0,"format":fmt_red})
    # Сводка категории
    wd=wb.add_worksheet("Сводка_Категории")
    wd.write(0,0,"СВОДКА ПО КАТЕГОРИЯМ (FBS)", wb.add_format({"bold":True,"font_size":14,"font_color":"#0F3460"}))
    wd.write_row(2,0,["Категория","Кол-во SKU","Выручка","Расходы","Прибыль","Маржа %","Убыточных"], fmt_h)
    cats=df["Категория"].dropna().unique().tolist()
    for i,cat in enumerate(cats):
        r=3+i
        wd.write(r,0,cat)
        wd.write_formula(r,1, f"=COUNTIF('Расчет_FBS'!C:C,A{r+1})")
        wd.write_formula(r,2, f"=SUMIF('Расчет_FBS'!C:C,A{r+1},'Расчет_FBS'!D:D)")
        wd.write_formula(r,3, f"=SUMIF('Расчет_FBS'!C:C,A{r+1},'Расчет_FBS'!K:K)")
        wd.write_formula(r,4, f"=SUMIF('Расчет_FBS'!C:C,A{r+1},'Расчет_FBS'!L:L)")
        wd.write_formula(r,5, f"=IF(C{r+1}>0,E{r+1}/C{r+1},0)")
        wd.write_formula(r,6, f"=COUNTIFS('Расчет_FBS'!C:C,A{r+1},'Расчет_FBS'!L:L,\"<0\")")
    wd.set_column(2,4,16,fmt_m); wd.set_column(5,5,12,fmt_p)
    # Сводка бренды
    wb2=wb.add_worksheet("Сводка_Бренды")
    wb2.write_row(2,0,["Бренд","Кол-во SKU","Выручка","Расходы","Прибыль","Маржа %","Убыточных"], fmt_h)
    brands=df["Бренд"].dropna().unique().tolist()
    for i,br in enumerate(brands):
        r=3+i
        wb2.write(r,0,br)
        wb2.write_formula(r,1, f"=COUNTIF('Расчет_FBS'!B:B,A{r+1})")
        wb2.write_formula(r,2, f"=SUMIF('Расчет_FBS'!B:B,A{r+1},'Расчет_FBS'!D:D)")
        wb2.write_formula(r,3, f"=SUMIF('Расчет_FBS'!B:B,A{r+1},'Расчет_FBS'!K:K)")
        wb2.write_formula(r,4, f"=SUMIF('Расчет_FBS'!B:B,A{r+1},'Расчет_FBS'!L:L)")
        wb2.write_formula(r,5, f"=IF(C{r+1}>0,E{r+1}/C{r+1},0)")
        wb2.write_formula(r,6, f"=COUNTIFS('Расчет_FBS'!B:B,A{r+1},'Расчет_FBS'!L:L,\"<0\")")
    wb2.set_column(2,4,16,fmt_m); wb2.set_column(5,5,12,fmt_p)
    # Легенда
    wl=wb.add_worksheet("Легенда")
    leg=[["Колонка","Описание"],
         ["Артикул","Уникальный идентификатор товара"],
         ["Бренд","Производитель — разрез аналитики"],
         ["Категория","Группа товаров: габариты и спецтарифы"],
         ["Длина/Ширина/Высота","Габариты упаковки, см"],
         ["Объем_л","Д×Ш×В ÷ 1000 или из файла/справочника"],
         ["Оплач_вес","max(вес; объёмный вес Д×Ш×В÷5000), мин 0.1 кг"],
         ["Цена","Розничная цена на Маркете"],
         ["Себестоимость","Из файла или % от цены"],
         ["Комиссия_руб","MAX(Цена×ставка; мин. комиссия)"],
         ["Логистика_руб","База + Оплач.вес×ставка"],
         ["Хранение_руб","Объём×ставка×оборачиваемость"],
         ["Эквайринг_руб","Цена×ставка эквайринга"],
         ["Спец_расходы_FBS","Упаковка, маркировка, резервы"],
         ["Итого_расходы","Сумма всех расходов"],
         ["Прибыль","Цена − Итого_расходы"],
         ["Маржа_%","Прибыль ÷ Цена"],
         ["Спецтариф_применён","Флаг спецтарифа"],
         ["Причина_спецтарифа","Пояснение"]]
    for i,row in enumerate(leg):
        if i==0: wl.write_row(i,0,row,fmt_h)
        else: wl.write_row(i,0,row)
    wl.set_column(0,0,22); wl.set_column(1,1,64)
    wb.close()
    return out.getvalue()
def build_excel_values(df: pd.DataFrame, tariff:dict) -> bytes:
    out=io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter") as w:
        wb=w.book
        fmt_h=wb.add_format({"bold":True,"bg_color":"#0F3460","font_color":"#FFFFFF","border":1,"align":"center"})
        fmt_m=wb.add_format({"num_format":"#,##0.00","border":1})
        fmt_p=wb.add_format({"num_format":"0.00%","border":1})
        # Итоги
        t=totals_row(df)
        ws=w.book.add_worksheet("Итоги")
        ws.write_row(0,0,["Показатель","Значение"],fmt_h)
        for i,(k,v) in enumerate([("SKU",t["count"]),("Выручка, ₽",t["revenue"]),("Расходы, ₽",t["expenses"]),("Прибыль, ₽",t["profit"]),("Маржа",t["avg_margin"]),("Убыточных SKU",t["loss"])], start=1):
            ws.write(i,0,k); ws.write(i,1,v, fmt_p if k=="Маржа" else fmt_m if isinstance(v,float) else None)
        ws.set_column(0,0,22); ws.set_column(1,1,16)
        # Расчёт (значения)
        cols=["Артикул","Бренд","Категория","Длина","Ширина","Высота","Объем_л","Оплач_вес","Цена","Себестоимость","Комиссия_руб","Логистика_руб","Хранение_руб","Эквайринг_руб","Спец_расходы_FBS","Итого_расходы","Прибыль","Маржа_%","Спецтариф_применён","Причина_спецтарифа"]
        df[cols].to_excel(w, sheet_name="Расчет_FBS", index=False)
        ws2=w.sheets["Расчет_FBS"]
        for c in range(len(cols)): ws2.write(0,c,cols[c],fmt_h)
        ws2.freeze_panes(1,0); ws2.autofilter(0,0,len(df),len(cols)-1)
        # Сводки
        for name,by in [("Сводка_Категории","Категория"),("Сводка_Бренды","Бренд")]:
            g=summarize(df, by)
            ws3=wb.add_worksheet(name)
            ws3.write_row(0,0,[by,"Кол-во SKU","Выручка","Расходы","Прибыль","Маржа %","Убыточных","Спецтариф"],fmt_h)
            for i,(_,r) in enumerate(g.iterrows(), start=1):
                ws3.write(i,0,r["name"]); ws3.write(i,1,int(r["count"])); ws3.write(i,2,float(r["revenue"]),fmt_m)
                ws3.write(i,3,float(r["expenses"]),fmt_m); ws3.write(i,4,float(r["profit"]),fmt_m)
                ws3.write(i,5,float(r["avg_margin"]),fmt_p); ws3.write(i,6,int(r["loss"])); ws3.write(i,7,int(r["spec"]))
            ws3.set_column(0,0,28); ws3.set_column(2,4,16,fmt_m); ws3.set_column(5,5,12,fmt_p)
        # Тариф
        ws4=wb.add_worksheet("Тариф")
        ws4.write_row(0,0,["Параметр","Значение"],fmt_h)
        for i,(k,v) in enumerate([("Комиссия",tariff["base"]["commission_rate"]),("Мин. комиссия",tariff["base"]["min_commission"]),("Логистика база",tariff["base"]["logistics_base"]),("Логистика за кг",tariff["base"]["logistics_per_kg"]),("Хранение за л/сутки",tariff["base"]["storage_per_day_per_liter"]),("Эквайринг",tariff["base"]["acquiring_fee"])], start=1):
            ws4.write(i,0,k); ws4.write(i,1,v, fmt_p if "%" in k or "Комиссия"==k else fmt_m)
    return out.getvalue()
# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT APP
# ─────────────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Юнит-экономика FBS — Яндекс Маркет", layout="wide", page_icon="🚗")
st.markdown(CSS, unsafe_allow_html=True)
if "tariff" not in st.session_state: st.session_state.tariff=default_tariff()
if "df_calc" not in st.session_state: st.session_state.df_calc=None
if "df_raw" not in st.session_state: st.session_state.df_raw=None
if "parse_info" not in st.session_state: st.session_state.parse_info=None
if "step" not in st.session_state: st.session_state.step=0
tariff=st.session_state.tariff
# HERO
st.markdown(f"""
<div class="hero">
  <div style="display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap;align-items:center;position:relative">
    <div>
      <h1>{APP_NAME} <span style="color:#fbbf24">· FBS</span></h1>
      <p>Автозапчасти · До 300 000 SKU · Excel с живыми формулами · v{APP_VERSION}</p>
    </div>
    <div style="display:flex;gap:8px">
      <span class="badge">⚡ Потоковый CSV</span>
      <span class="badge">📊 Сводки по брендам</span>
      <span class="badge">🧮 Виртуализация 300k</span>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)
# STEPPER
steps=[("⚙️","Тарифы","Комиссии и расходы"),("📦","Данные","Загрузите каталог"),("📊","Дашборд","Метрики и графики"),("📥","Экспорт","Excel и CSV")]
cols=st.columns(4)
for i,(icon,label,desc) in enumerate(steps):
    active = st.session_state.step==i
    done = st.session_state.df_calc is not None and i < 3
    max_reached = 1 if st.session_state.df_raw is None else 3 if st.session_state.df_calc is not None else 2
    clickable = i <= max_reached
    with cols[i]:
        if st.button(f"{icon} {label}\n{desc}", key=f"step{i}", use_container_width=True,
                     disabled=not clickable,
                     type="primary" if active else "secondary"):
            st.session_state.step=i
            st.rerun()
st.divider()
step=st.session_state.step
# ══════════════════════════════════════════════════════════════════════════════
# ШАГ 0 — ТАРИФЫ
# ══════════════════════════════════════════════════════════════════════════════
if step==0:
    st.info("💡 **Новичок?** Ничего не меняйте — тарифы уже заполнены типовыми ставками Яндекс Маркета. Переходите к шагу «Данные».")
    c1,c2 = st.columns(2)
    with c1:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("⚙️ Базовый тариф Яндекс Маркета")
        b=tariff["base"]
        b["commission_rate"]=st.number_input("Комиссия маркетплейса, %", 0.0, 40.0, b["commission_rate"]*100, 0.5, help="Процент с продажи")/100
        b["min_commission"]=st.number_input("Минимальная комиссия, ₽", 0.0, 1000.0, b["min_commission"], 5.0)
        b["logistics_base"]=st.number_input("Логистика: базовая ставка, ₽", 0.0, 1000.0, b["logistics_base"], 5.0)
        b["logistics_per_kg"]=st.number_input("Логистика: за 1 кг, ₽", 0.0, 200.0, b["logistics_per_kg"], 0.5)
        b["storage_per_day_per_liter"]=st.number_input("Хранение за 1 л/сутки, ₽", 0.0, 5.0, b["storage_per_day_per_liter"], 0.05, format="%.2f")
        b["acquiring_fee"]=st.number_input("Эквайринг, %", 0.0, 10.0, b["acquiring_fee"]*100, 0.1)/100
        b["return_fee"]=st.number_input("Возвраты, %", 0.0, 10.0, b["return_fee"]*100, 0.1)/100
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.subheader("🔧 Спецтарифы автозапчастей")
        tariff["special_enabled"]=st.toggle("Применять спецтарифы", value=tariff["special_enabled"],
                                            help="Шины, АКБ, двигатели и КПП — отдельные ставки по ключевым словам категории")
        for key, rule in SPECIAL_TARIFFS.items():
            with st.expander(f"{rule['label']}  —  ключ «{key}»", expanded=False):
                cA,cB,cC=st.columns(3)
                # храним переопределения в tariff, иначе берём дефолт
                if key not in tariff: tariff[key]=dict(rule)
                cur=tariff.get(key, rule)
                cur["commission_rate"]=cA.number_input(f"Комиссия {rule['label']}, %", 0.0,40.0, cur["commission_rate"]*100,0.5, key=f"sp_c_{key}")/100
                cur["logistics_base"]=cB.number_input(f"Логистика {rule['label']}, ₽", 0.0,500.0, cur["logistics_base"],5.0, key=f"sp_l_{key}")
                cur["storage_per_day_per_liter"]=cC.number_input(f"Хранение {rule['label']}, ₽/л", 0.0,5.0, cur["storage_per_day_per_liter"],0.05, key=f"sp_s_{key}")
                tariff[key]=cur
                # синхронизируем SPECIAL_TARIFFS для расчёта
                SPECIAL_TARIFFS[key].update(cur)
        st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div class="card" style="margin-top:16px">', unsafe_allow_html=True)
    st.subheader("🧩 Если данных нет в файле")
    sc=tariff["special_costs"]
    a,b,c = st.columns(3)
    sc["packaging"]=a.number_input("Упаковка FBS, ₽", 0.0,500.0, sc["packaging"],5.0)
    sc["chestny_znak"]=b.number_input("Честный знак, ₽", 0.0,50.0, sc["chestny_znak"],0.5)
    sc["labeling"]=c.number_input("Маркировка, ₽", 0.0,50.0, sc["labeling"],0.5)
    a,b,c = st.columns(3)
    sc["warranty_reserve"]=a.number_input("Гарант. резерв, %", 0.0,10.0, sc["warranty_reserve"]*100,0.5)/100
    sc["hazard_surcharge"]=b.number_input("Надбавка опасный, %", 0.0,10.0, sc["hazard_surcharge"]*100,0.5)/100
    sc["fragile_surcharge"]=c.number_input("Надбавка хрупкий, %", 0.0,10.0, sc["fragile_surcharge"]*100,0.5)/100
    x,y = st.columns(2)
    tariff["cost_fallback"]=x.slider("Себестоимость по умолчанию, % от цены", 10,95,int(tariff["cost_fallback"]*100))/100
    tariff["density"]=y.number_input("Плотность для оценки веса, кг/л", 0.05,2.0, tariff["density"],0.05)
    st.caption("Если колонки «Себестоимость» нет — она рассчитается как этот процент от цены. Если веса нет — оценим как объём × плотность.")
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div class="card" style="margin-top:16px">', unsafe_allow_html=True)
    st.subheader("🌐 Индивидуальные ставки категорий (API)")
    st.caption("Точные комиссии по категориям из партнёрского API Яндекс Маркета")
    t1,t2,t3=st.columns([2,1,1])
    token=t1.text_input("OAuth-токен", type="password", placeholder="Введите токен…")
    campaign=t2.number_input("Campaign ID", min_value=1, value=123456, step=1)
    if t3.button("Обновить из API", use_container_width=True):
        if not token:
            st.warning("Введите OAuth-токен")
        else:
            try:
                r=requests.get(f"https://api.partner.market.yandex.ru/v2/campaigns/{int(campaign)}/categories/commissions",
                               headers={"Authorization":f"OAuth {token}","Accept":"application/json"}, timeout=12)
                if r.status_code!=200:
                    st.error(f"API вернул {r.status_code}: {r.text[:300]}")
                else:
                    data=r.json(); cats=data.get("result",{}).get("categories",[])
                    if not cats: st.warning("API вернул пустой список категорий")
                    else:
                        rates={}
                        for c in cats:
                            name=(c.get("categoryName") or "").lower()
                            pct_=c.get("commissionPercent",14)/100
                            if name: rates[name]=pct_
                        tariff["category_rates"]=rates
                        tariff["use_category_rates"]=True
                        st.success(f"Получены ставки для {len(cats)} категорий")
            except Exception as e:
                st.warning(f"Прямой запрос из браузера/Streamlit Cloud может блокироваться CORS. Ошибка: {e}. Используйте демо-ставки ниже.")
    if st.button("🎲 Загрузить демо-ставки категорий"):
        tariff["category_rates"]={"фильтры":0.14,"масла":0.15,"колодки":0.13,"диски":0.16,"амортизаторы":0.15,"аккумуляторы":0.13,"шины":0.12,"фары":0.17,"двигатели":0.11,"кпп":0.11}
        tariff["use_category_rates"]=True
        st.success("Демо-ставки загружены")
    if tariff.get("category_rates"):
        st.write("Загружено ставок:", len(tariff["category_rates"]))
        st.json(tariff["category_rates"])
        tariff["use_category_rates"]=st.toggle("Использовать индивидуальные ставки", value=tariff["use_category_rates"])
    st.markdown('</div>', unsafe_allow_html=True)
    # Скачать текущий тариф CSV
    csv_lines=["Параметр;Значение",
               f"Комиссия, %;{tariff['base']['commission_rate']*100:.2f}".replace(".",","),
               f"Мин. комиссия, ₽;{tariff['base']['min_commission']}",
               f"Логистика база, ₽;{tariff['base']['logistics_base']}",
               f"Логистика за кг, ₽;{tariff['base']['logistics_per_kg']}",
               f"Хранение за л/сутки, ₽;{tariff['base']['storage_per_day_per_liter']}",
               f"Эквайринг, %;{tariff['base']['acquiring_fee']*100:.2f}".replace(".",",")]
    st.download_button("⬇️ Скачать текущий тариф (CSV)", data=("\ufeff"+"\n".join(csv_lines)).encode("utf-8-sig"),
                       file_name="current_tariff.csv", mime="text/csv")
# ══════════════════════════════════════════════════════════════════════════════
# ШАГ 1 — ДАННЫЕ
# ══════════════════════════════════════════════════════════════════════════════
elif step==1:
    st.markdown('<div class="card">', unsafe_allow_html=True)
    st.subheader("📦 Загрузка каталога")
    st.caption("CSV или Excel · до 300 000 SKU · Обязательные колонки: **Артикул, Категория, Цена**.  Бренд, габариты, себестоимость — необязательны.")
    up=st.file_uploader("Перетащите файл сюда или нажмите «Browse files»", type=["csv","xlsx","xls","txt","tsv"])
    st.markdown('</div>', unsafe_allow_html=True)
    cA,cB,cC,cD=st.columns(4)
    if cA.button("📋 Шаблон CSV", use_container_width=True):
        tpl="Артикул;Бренд;Категория;Длина;Ширина;Высота;Цена;Себестоимость;Вес_кг;Оборачиваемость_дней\nMAN-FLT-000001;Mann-Filter;Фильтры;22;14;14;450;220;0,5;25\nMIC-TR-000002;Michelin;Шины;70;70;26;5400;3650;10,5;15\nVAR-BAT-000003;Varta;Аккумуляторы;35;26;26;6500;;16,5;30\n"
        st.download_button("Скачать template_fbs.csv", data=("\ufeff"+tpl).encode("utf-8-sig"), file_name="template_fbs.csv", mime="text/csv", key="tpl_dl")
    if cB.button("🎲 Пример 24 SKU", use_container_width=True):
        df_demo=build_demo_rows(24)
        st.session_state.df_raw, _,_,_,_,_,_=parse_dataframe(df_demo, "demo_24.csv", 9999)
        # сохраняем как raw для парсинга — но уже есть df_raw, просто пересобираем parse_info
        st.session_state.parse_info=dict(file_name="demo_24.csv", file_size=9999, matched={"Артикул":"Артикул","Бренд":"Бренд","Категория":"Категория","Цена":"Цена","Себестоимость":"Себестоимость","Длина":"Длина","Ширина":"Ширина","Высота":"Высота"}, missing=[], has_cost=True, has_dims=True, skipped=0, parse_ms=0, rows=len(df_demo))
        # триггер расчёта
        st.session_state.df_calc=calculate_df(st.session_state.df_raw, st.session_state.tariff)
        st.success(f"Загружено демо: {len(df_demo)} SKU"); st.rerun()
    if cC.button("⚡ 50 000 SKU", use_container_width=True):
        with st.spinner("Генерируем 50 000 SKU…"):
            df_demo=build_demo_rows(50_000)
            st.session_state.df_raw,_,_,_,_,_,_=parse_dataframe(df_demo, "demo_50000.csv", 50000*120)
            st.session_state.parse_info=dict(file_name="demo_50000.csv", file_size=50000*120, matched={}, missing=[], has_cost=True, has_dims=True, skipped=0, parse_ms=0, rows=len(df_demo))
            st.session_state.df_calc=calculate_df(st.session_state.df_raw, st.session_state.tariff)
        st.success("Готово: 50 000 SKU"); st.rerun()
    if cD.button("🚀 300 000 SKU", use_container_width=True):
        with st.spinner("Генерируем 300 000 SKU — это займёт ~5 с…"):
            df_demo=build_demo_rows(300_000)
            st.session_state.df_raw,_,_,_,_,_,_=parse_dataframe(df_demo, "demo_300000.csv", 300000*120)
            st.session_state.parse_info=dict(file_name="demo_300000.csv", file_size=300000*120, matched={}, missing=[], has_cost=True, has_dims=True, skipped=0, parse_ms=0, rows=len(df_demo))
            # чанками
            prog=st.progress(0, text="Расчёт 300 000 SKU…")
            # calculate_df уже векторизован и быстрый — один вызов
            st.session_state.df_calc=calculate_df(st.session_state.df_raw, st.session_state.tariff)
            prog.progress(100, text="Готово")
        st.success("Готово: 300 000 SKU"); st.rerun()
    if up is not None:
        try:
            raw_bytes=up.getvalue()
            t0=time.time()
            if up.name.lower().endswith(".csv") or up.name.lower().endswith(".txt") or up.name.lower().endswith(".tsv"):
                # пробуем ; затем ,
                try: df_raw=pd.read_csv(io.BytesIO(raw_bytes), sep=";", dtype=str, keep_default_na=False)
                except: df_raw=pd.read_csv(io.BytesIO(raw_bytes), sep=",", dtype=str, keep_default_na=False)
                # если одна колонка — пробуем другой разделитель
                if df_raw.shape[1]==1:
                    try: df_raw=pd.read_csv(io.BytesIO(raw_bytes), sep=",", dtype=str, keep_default_na=False)
                    except: pass
            else:
                if len(raw_bytes) > 60*1024*1024:
                    st.error("Excel >60 МБ — сохраните в CSV для потокового чтения"); st.stop()
                df_raw=pd.read_excel(io.BytesIO(raw_bytes), dtype=str, keep_default_na=False)
            df_parsed, cmap, missing, has_cost, has_dims, skipped, parse_ms = parse_dataframe(df_raw, up.name, len(raw_bytes))
            st.session_state.df_raw=df_parsed
            st.session_state.parse_info=dict(file_name=up.name, file_size=len(raw_bytes), matched={k:v for k,v in cmap.items() if v}, missing=missing, has_cost=has_cost, has_dims=has_dims, skipped=skipped, parse_ms=parse_ms, rows=len(df_parsed))
            # расчёт
            prog_text="Расчёт…" if len(df_parsed)<BIG_DATA_THRESHOLD else f"Расчёт {len(df_parsed):,} SKU — не закрывайте вкладку…"
            with st.spinner(prog_text):
                df_calc=calculate_df(df_parsed, st.session_state.tariff)
            st.session_state.df_calc=df_calc
            st.success(f"Загружено {len(df_parsed):,} SKU за {parse_ms/1000:.1f} с · расчёт за {(time.time()-t0):.1f} с")
            st.rerun()
        except Exception as e:
            logger.exception(e)
            st.error(f"Ошибка обработки файла: {e}")
    # предпросмотр
    if st.session_state.parse_info is not None:
        info=st.session_state.parse_info
        st.markdown('<div class="card" style="margin-top:16px">', unsafe_allow_html=True)
        st.subheader(f"🔍 Предпросмотр: {info['file_name']}")
        c1,c2,c3=st.columns(3)
        c1.metric("Строк", f"{info['rows']:,}".replace(","," "))
        c2.metric("Размер", fmt_size(info["file_size"]))
        c3.metric("Пустых пропущено", info["skipped"])
        if info["missing"]: st.error(f"Не найдены обязательные колонки: {', '.join(info['missing'])}")
        if not info["has_cost"]: st.warning("Колонка «Себестоимость» не найдена — будет рассчитана как % от цены (настраивается в тарифах). Значения с «~» в таблице.")
        if not info["has_dims"]: st.info("Габариты не найдены — объём и вес возьмём из справочника категорий.")
        if info["matched"]:
            st.caption("Сопоставленные колонки: " + " · ".join(f"{k} ← {v}" for k,v in info["matched"].items()))
        if st.session_state.df_raw is not None:
            st.dataframe(st.session_state.df_raw.head(8), use_container_width=True, hide_index=True)
            if info["rows"]>8: st.caption(f"…и ещё {(info['rows']-8):,} строк. Полная таблица — на дашборде.")
        st.markdown('</div>', unsafe_allow_html=True)
        if st.session_state.df_calc is not None and st.button("Смотреть дашборд →", type="primary"):
            st.session_state.step=2; st.rerun()
# ══════════════════════════════════════════════════════════════════════════════
# ШАГ 2 — ДАШБОРД
# ══════════════════════════════════════════════════════════════════════════════
elif step==2:
    df=st.session_state.df_calc
    if df is None or df.empty:
        st.warning("Сначала загрузите каталог на шаге «Данные»")
        if st.button("К загрузке данных"): st.session_state.step=1; st.rerun()
        st.stop()
    t=totals_row(df)
    by_cat=summarize(df,"Категория")
    by_brand=summarize(df,"Бренд")
    # метрики
    m1,m2,m3,m4=st.columns(4)
    for col, title, val, sub, tone in [
        (m1,"Выручка", money(t["revenue"]), f"{t['count']:,} SKU · {len(by_brand)} брендов","mc-indigo"),
        (m2,"Итого расходы", money(t["expenses"]), f"Доля расходов {pct(t['expenses']/t['revenue'] if t['revenue'] else 0,0)}","mc-violet"),
        (m3,"Чистая прибыль", money(t["profit"]), "После всех комиссий" if t["profit"]>=0 else "Убыток — поднимите цены","mc-rose" if t["profit"]<0 else "mc-emerald"),
        (m4,"Маржа по обороту", pct(t["avg_margin"]), f"Убыточных SKU: {t['loss']:,}","mc-emerald" if t["avg_margin"]>=0.15 else "mc-amber" if t["avg_margin"]>=0 else "mc-rose"),
    ]:
        with col:
            st.markdown(f'<div class="metric-card {tone}"><div class="small">{title}</div><div style="font-size:1.55rem;font-weight:800;margin-top:4px">{val}</div><div class="small" style="margin-top:4px">{sub}</div></div>', unsafe_allow_html=True)
    s1,s2,s3,s4,s5=st.columns(5)
    for col, icon, label, val in [
        (s1,"🏷️","Эфф. комиссия", pct(t["commission"]/t["revenue"] if t["revenue"] else 0,1)),
        (s2,"📦","Ср. логистика", money(t["logistics"]/t["count"] if t["count"] else 0)),
        (s3,"🏬","Ср. хранение", money(t["storage"]/t["count"] if t["count"] else 0)),
        (s4,"🔧","Спецтариф", f"{t['spec']:,} SKU"),
        (s5,"~","Себест. оценена", f"{t['est']:,} SKU"),
    ]:
        with col: st.markdown(f'<div class="card" style="text-align:center"><div style="font-size:1.2rem">{icon}</div><div class="small">{label}</div><div style="font-weight:700">{val}</div></div>', unsafe_allow_html=True)
    # переключатель разреза
    st.write("")
    mode=st.radio("Разрез аналитики", ["📂 По категориям","🏭 По брендам"], horizontal=True, label_visibility="collapsed")
    by_mode = "Категория" if "категори" in mode.lower() else "Бренд"
    g = by_cat if by_mode=="Категория" else by_brand
    st.caption(f"Всего групп: {len(g):,} · на графиках топ-12 по выручке")
    # графики
    top12=g.head(12)
    # маржа
    fig1=px.bar(top12, x="name", y="avg_margin", color="avg_margin", color_continuous_scale="RdYlGn", range_color=[-0.2,0.35],
                labels={"name":by_mode,"avg_margin":"Маржа"}, title=f"Маржа по {by_mode.lower()}м")
    fig1.update_layout(height=360, margin=dict(l=10,r=10,t=40,b=80), coloraxis_showscale=False,
                       xaxis_tickangle=-22, yaxis_tickformat=".0%")
    # структура расходов
    costs_df=pd.DataFrame([("Себестоимость",t["cost"],"#6366f1"),("Комиссия",t["commission"],"#f59e0b"),("Логистика",t["logistics"],"#0ea5e9"),("Хранение",t["storage"],"#8b5cf6"),("Эквайринг",float(df["Эквайринг_руб"].sum()),"#14b8a6"),("Спец. расходы",float(df["Спец_расходы_FBS"].sum()),"#f43f5e")], columns=["name","value","color"])
    costs_df=costs_df[costs_df["value"]>0]
    fig2=px.pie(costs_df, names="name", values="value", hole=0.58, title="Структура расходов")
    fig2.update_traces(marker=dict(colors=costs_df["color"]), textinfo="percent+label")
    fig2.update_layout(height=360, margin=dict(l=10,r=10,t=40,b=10))
    # прибыль
    fig3=px.bar(top12, x="name", y="profit", color=top12["profit"]>=0, color_discrete_map={True:"#10b981", False:"#ef4444"},
                labels={"name":by_mode,"profit":"Прибыль"}, title=f"Прибыль по {by_mode.lower()}м")
    fig3.update_layout(height=360, margin=dict(l=10,r=10,t=40,b=80), showlegend=False, xaxis_tickangle=-22,
                       yaxis_tickformat=",.0f")
    a,b=st.columns(2)
    with a: st.plotly_chart(fig1, use_container_width=True)
    with b: st.plotly_chart(fig2, use_container_width=True)
    st.plotly_chart(fig3, use_container_width=True)
    # сводная таблица групп
    st.subheader(f"Сводка по {by_mode.lower()}м — топ-20 по выручке")
    show_cols=["name","count","revenue","expenses","profit","avg_margin","loss"]
    rename={"name":by_mode,"count":"SKU","revenue":"Выручка","expenses":"Расходы","profit":"Прибыль","avg_margin":"Маржа","loss":"Убыточных"}
    tbl=g.head(20).rename(columns=rename)
    st.dataframe(tbl.style.format({"Выручка":"{:,.0f} ₽","Расходы":"{:,.0f} ₽","Прибыль":"{:,.0f} ₽","Маржа":"{:.1%}"}),
                 use_container_width=True, hide_index=True)
    # убыточные SKU
    losses=df[df["Прибыль"]<0].sort_values("Прибыль").head(8)
    if not losses.empty:
        st.subheader(f"🚨 Самые убыточные SKU — всего {t['loss']:,}")
        st.dataframe(losses[["Артикул","Бренд","Категория","Цена","Прибыль","Маржа_%"]].style.format({"Цена":"{:,.0f} ₽","Прибыль":"{:,.0f} ₽","Маржа_%":"{:.1%}"}),
                     use_container_width=True, hide_index=True)
    # полная таблица с фильтрами
    st.subheader("🧮 Расчёт по всем SKU")
    f1,f2,f3,f4=st.columns([2,1,1,1])
    q=f1.text_input("Поиск по артикулу / бренду / категории", placeholder="Начните вводить…")
    brands_sorted=df["Бренд"].value_counts().head(200).index.tolist()
    cats_sorted=df["Категория"].value_counts().head(200).index.tolist()
    br=f2.selectbox("Бренд", ["Все бренды"]+brands_sorted)
    cat=f3.selectbox("Категория", ["Все категории"]+cats_sorted)
    only_loss=f4.toggle("Только убыточные")
    view=df
    if q: view=view[view["Артикул"].str.contains(q,case=False, na=False) | view["Бренд"].str.contains(q,case=False, na=False) | view["Категория"].str.contains(q,case=False, na=False)]
    if br!="Все бренды": view=view[view["Бренд"]==br]
    if cat!="Все категории": view=view[view["Категория"]==cat]
    if only_loss: view=view[view["Прибыль"]<0]
    # сортировка
    sort_col=st.selectbox("Сортировка", ["Прибыль","Цена","Маржа_%","Комиссия_руб","Логистика_руб","Артикул","Бренд","Категория"], index=0)
    asc=st.toggle("По возрастанию", value=False)
    view=view.sort_values(sort_col, ascending=asc)
    st.caption(f"Показано {len(view):,} из {len(df):,} SKU")
    st.dataframe(view.head(5000) if len(view)>5000 else view, use_container_width=True, hide_index=True,
                 height=560)
    if len(view)>5000: st.info("Показаны первые 5000 строк — полный объём выгрузите в Excel/CSV на шаге «Экспорт».")
    if st.button("📥 К экспорту →", type="primary"): st.session_state.step=3; st.rerun()
# ══════════════════════════════════════════════════════════════════════════════
# ШАГ 3 — ЭКСПОРТ
# ══════════════════════════════════════════════════════════════════════════════
else:
    df=st.session_state.df_calc
    if df is None or df.empty:
        st.warning("Сначала загрузите каталог")
        if st.button("К загрузке данных"): st.session_state.step=1; st.rerun()
        st.stop()
    t=totals_row(df); n=len(df)
    st.markdown(f'<div class="card" style="background:linear-gradient(135deg,#059669,#0d9488);color:white;border:none"><h3 style="margin:0;color:white">Отчёт готов 🎉</h3><p style="margin:6px 0 0;color:#d1fae5">{n:,} SKU · Выручка {money(t["revenue"])} · Прибыль {money(t["profit"])} · Маржа {pct(t["avg_margin"])}</p></div>', unsafe_allow_html=True)
    st.write("")
    c1,c2,c3=st.columns(3)
    with c1:
        st.markdown('<div class="card" style="text-align:center">', unsafe_allow_html=True)
        st.write("**📊 Excel с живыми формулами**")
        st.caption(f"Листы: Тариф, Входные_Данные, Расчет_FBS (формулы), Сводка_Категории, Сводка_Бренды, Легенда. Лимит {FORMULA_ROW_LIMIT:,} строк.")
        if n>FORMULA_ROW_LIMIT:
            st.warning(f"Слишком много строк ({n:,}) — используйте быстрый Excel или CSV")
        else:
            if st.button("Скачать .xlsx с формулами", type="primary", use_container_width=True):
                with st.spinner("Формируем Excel с формулами…"):
                    data=build_excel_formula(df, st.session_state.tariff)
                st.download_button("⬇️ Скачать файл", data=data, file_name=f"unit_economy_fbs_formulas_{datetime.now():%Y%m%d_%H%M%S}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_formula")
        st.markdown('</div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="card" style="text-align:center">', unsafe_allow_html=True)
        st.write("**⚡ Быстрый Excel (значения)**")
        st.caption("Итоги, полный расчёт, сводки по категориям и брендам. Держит 300k строк.")
        if st.button("Скачать быстрый .xlsx", use_container_width=True):
            with st.spinner("Формируем быстрый Excel…"):
                data=build_excel_values(df, st.session_state.tariff)
            st.download_button("⬇️ Скачать файл", data=data, file_name=f"unit_economy_fbs_{datetime.now():%Y%m%d_%H%M%S}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="dl_values")
        st.markdown('</div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="card" style="text-align:center">', unsafe_allow_html=True)
        st.write("**📄 CSV со всеми SKU**")
        st.caption("Разделитель «;», UTF-8 BOM. Рекомендуется для 300k.")
        csv_buf=io.StringIO(); df.to_csv(csv_buf, index=False, sep=";", encoding="utf-8-sig")
        st.download_button("Скачать .csv", data=csv_buf.getvalue().encode("utf-8-sig"), file_name=f"unit_economy_fbs_{datetime.now():%Y%m%d_%H%M%S}.csv", mime="text/csv", use_container_width=True)
        st.markdown('</div>', unsafe_allow_html=True)
    st.divider()
    st.caption("Все расчёты выполняются локально. Данные не покидают ваш сервер.")
# footer
st.markdown(f"<p style='text-align:center;color:#94a3b8;font-size:.78rem;margin-top:18px'>{APP_NAME} · v{APP_VERSION} · монолит Streamlit</p>", unsafe_allow_html=True)
