"""
Кэш показов Яндекса по дням на диске (защита от нулевых «дыр» в выгрузке).

Отчёт shows-sales иногда возвращает 0 по уже бывшему дню (данные Яндекса
«схлопываются» в конкретном снимке, потом доуточняются). Это рушит лист
«Показы 14д» и анализ буст/органики.

Логика: держим снимок метрик по (день, SKU) в shows_cache.json.
- День НЕнулевой в живой выгрузке (shows>0) → перезаписываем кэш свежими данными.
- День нулевой/пустой, а в кэше было ненулевое → подставляем значение из кэша.
Так живые данные всегда главнее, но нулевая дыра не затирает ранее полученное.
"""

import json
import os
from datetime import date, timedelta

import pandas as pd

_CACHE_FILE = os.path.join(os.path.dirname(__file__), "shows_cache.json")

# Метрики, которые храним и восстанавливаем (берём те, что есть в df).
_METRIC_COLS = [
    "shows", "shows_promo", "shows_boost_pct",
    "clicks", "clicks_promo",
    "carts", "carts_promo", "carts_boost_pct",
    "orders", "orders_promo", "orders_boost_pct",
    "revenue", "revenue_promo",
    "visibility", "ctr_to_cart", "cart_to_order",
]


def _num(v):
    """В число (float) или None."""
    try:
        if v is None or (isinstance(v, float) and pd.isna(v)):
            return None
        return float(v)
    except (TypeError, ValueError):
        return None


def load_cache() -> dict:
    if not os.path.exists(_CACHE_FILE):
        return {}
    try:
        with open(_CACHE_FILE, encoding="utf-8") as f:
            return json.load(f)
    except (json.JSONDecodeError, OSError):
        return {}


def save_cache(cache: dict) -> None:
    with open(_CACHE_FILE, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False)


def merge_with_cache(df: pd.DataFrame, date_from: date, date_to: date) -> pd.DataFrame:
    """
    Слить живую выгрузку с кэшем. Возвращает df, в котором нулевые/отсутствующие
    дни заполнены из кэша. Кэш на диске обновляется ненулевыми днями.
    Требует колонки df: 'date' (datetime.date), 'shopSku'.
    """
    if df is None or df.empty or "shopSku" not in df.columns or "date" not in df.columns:
        return df

    cache = load_cache()
    cols = [c for c in _METRIC_COLS if c in df.columns]

    # 1. Обновляем кэш ненулевыми живыми строками; собираем «хорошие» ключи.
    live_good = set()
    for _, r in df.iterrows():
        sku = r.get("shopSku")
        if sku is None or (isinstance(sku, float) and pd.isna(sku)):
            continue
        shows = _num(r.get("shows"))
        ds = r["date"].isoformat()
        if shows and shows > 0:
            cache.setdefault(ds, {})[str(sku)] = {c: _num(r.get(c)) for c in cols}
            live_good.add((ds, str(sku)))

    # 2. Подставляем из кэша строки, которых нет/нулевые в живой выгрузке
    #    (только в пределах окна отчёта).
    inject = []
    inject_keys = set()
    for ds, skus in cache.items():
        try:
            d = date.fromisoformat(ds)
        except ValueError:
            continue
        if d < date_from or d > date_to:
            continue
        for sku, metrics in skus.items():
            if (ds, sku) in live_good:
                continue
            row = {"date": d, "shopSku": sku}
            row.update({k: v for k, v in metrics.items() if v is not None})
            inject.append(row)
            inject_keys.add((ds, sku))

    # Обрезка: дни старше 90 суток от конца окна больше не нужны
    # (иначе файл растёт бесконечно, ~+20КБ/день).
    cutoff = (date_to - timedelta(days=90)).isoformat()
    for ds in [k for k in cache if k < cutoff]:
        del cache[ds]

    save_cache(cache)

    if not inject:
        return df

    # Убираем из df заменяемые (нулевые) строки и дописываем кэшевые.
    key = df["date"].map(lambda x: x.isoformat()) + "|" + df["shopSku"].astype(str)
    drop_mask = key.isin({f"{ds}|{sku}" for ds, sku in inject_keys})
    df = pd.concat([df[~drop_mask], pd.DataFrame(inject)], ignore_index=True)
    return df
