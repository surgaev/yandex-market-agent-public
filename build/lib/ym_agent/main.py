#!/usr/bin/env python3
"""
Yandex Market daily analytics agent.

Периоды: вчера / позавчера / неделю назад (сегодня не берём — данные ещё копятся).
Фильтр: товар попадает в отчёт только если есть показы хотя бы за один из трёх дней.
Витрина: парсим публичную страницу магазина, отмечаем товары которых нет.
Буст: считаем по вчерашним данным.

Usage:
    YM_OAUTH_TOKEN=<token> YM_REPORT_FOLDER=/path python main.py [--date YYYY-MM-DD]
"""

from .config import DATA_DIR
import argparse
import statistics
import sys
import traceback
import pandas as pd
from datetime import date, timedelta


def _enable_timestamps():
    """Префиксует все print текущим временем [HH:MM:SS] — чтобы в agent.log
    было видно, какая фаза сколько заняла (прогоны бывали по 2-3.5 часа,
    и без таймстампов виновника не найти)."""
    import builtins
    from datetime import datetime as _dt
    orig = builtins.print
    def tprint(*args, **kw):
        if args:
            orig(f"[{_dt.now().strftime('%H:%M:%S')}]", *args, **kw)
        else:
            orig(**kw)
    builtins.print = tprint

from .config import REPORT_FOLDER, THRESHOLD_RATIO, EXCLUDED_SKUS
from .yandex_client import (get_all_offers, fetch_shows_sales, fetch_placement_fee_detail,
                           get_current_bids, fetch_netting)
from .storefront_parser import get_storefront_skus
from .fbs_fby_client import (get_orders_by_type, get_stocks, estimate_fby_restock,
                            get_orders_with_fees, get_catalog_prices, get_promo_prices,
                            PromoFetchPartial, get_price_recommendations, get_reviews_summary)
from .report_generator import build_report
from .shows_cache import merge_with_cache


def _load_boost_snapshot():
    """Читает последний снимок ставок буста из boost_bids_history.csv.
    Returns: ({sku: {"bid_pct": float|None, "campaign": str}}, дата_снимка|None)."""
    import csv, os
    path = os.path.join(DATA_DIR, "boost_bids_history.csv")
    if not os.path.exists(path):
        return {}, None
    rows = list(csv.DictReader(open(path, encoding="utf-8")))
    if not rows:
        return {}, None
    last_date = max(r["date"] for r in rows if r.get("date"))
    snap = {}
    for r in rows:
        if r.get("date") != last_date:
            continue
        try:
            bid = float(r["bid_pct"]) if r.get("bid_pct") else None
        except ValueError:
            bid = None
        def _f(key):
            try: return float(r[key]) if r.get(key) else None
            except ValueError: return None
        snap[r["sku"]] = {"bid_pct": bid, "campaign": r.get("campaign", ""),
                          "price": _f("price"), "showcase_price": _f("showcase_price")}
    return snap, last_date


def _record_daily_bids(run_date, api_bids_today: dict):
    """Пишет сегодняшние API-ставки в boost_bids_daily.csv (БЕЗ обрезки — храним всё)
    и возвращает (prev, history):
      prev    — ставки за предыдущую записанную дату {sku: bid_pct};
      history — вся накопленная история {date: {sku: bid_pct}}."""
    import csv, os
    path = os.path.join(DATA_DIR, "boost_bids_daily.csv")
    today = run_date.isoformat()
    rows = []
    if os.path.exists(path):
        rows = list(csv.DictReader(open(path, encoding="utf-8")))
    # предыдущая дата (до сегодня)
    prev_dates = sorted({r["date"] for r in rows if r.get("date") and r["date"] < today})
    prev = {}
    if prev_dates:
        last = prev_dates[-1]
        for r in rows:
            if r["date"] == last and r.get("bid_pct"):
                try: prev[r["sku"]] = float(r["bid_pct"])
                except ValueError: pass
    # дописываем сегодняшние (не дублируя сегодняшнюю дату)
    rows = [r for r in rows if r.get("date") != today]
    with open(path, "w", encoding="utf-8", newline="") as f:
        w = csv.writer(f)
        w.writerow(["date", "sku", "bid_pct"])
        for r in rows:
            w.writerow([r.get("date", ""), r.get("sku", ""), r.get("bid_pct", "")])
        for sku, bid in api_bids_today.items():
            w.writerow([today, sku, bid])
    # собираем полную историю (уже с сегодняшним днём)
    history = {}
    for r in rows:
        d = r.get("date"); sku = r.get("sku")
        if not d or not sku:
            continue
        try:
            history.setdefault(d, {})[sku] = float(r["bid_pct"]) if r.get("bid_pct") else None
        except ValueError:
            pass
    for sku, bid in api_bids_today.items():
        history.setdefault(today, {})[sku] = bid
    return prev, history


# Что реально пошло в отчёт — свежие данные или кэш (и какого возраста). Заполняется
# загрузчиками по ходу прогона, выводится строкой на дашборде: файл может собраться
# успешно и выглядеть нормально, будучи собранным на позавчерашних данных.
DATA_HEALTH: dict[str, str] = {}


def _drop_excluded(obj, skus):
    """Убирает снятые с продажи артикулы из любых входных данных отчёта.

    Один фильтр на выходе вместо правки каждого листа: артикул сидит в двух
    десятках структур (словари по SKU, по (SKU, модель), по дате→SKU, списки
    строк заказов, таблицы pandas), и пропустить хоть одну — значит оставить
    его висеть на каком-нибудь листе.
      • dict: выкидываем ключ = SKU или кортеж (SKU, …); остальное — вглубь;
      • list/set: выкидываем строки-словари с полем артикула, строки-SKU;
      • DataFrame: по колонке shopSku.
    """
    if not skus:
        return obj
    if isinstance(obj, pd.DataFrame):
        return obj[~obj["shopSku"].isin(skus)] if "shopSku" in obj.columns else obj
    if isinstance(obj, dict):
        out = {}
        for k, v in obj.items():
            if k in skus or (isinstance(k, tuple) and k and k[0] in skus):
                continue
            out[k] = _drop_excluded(v, skus)
        return out
    if isinstance(obj, (list, set)):
        keep = []
        for x in obj:
            if isinstance(x, str) and x in skus:
                continue
            if isinstance(x, dict) and any(x.get(f) in skus
                                           for f in ("shopSku", "sku", "offerId", "offer_id")):
                continue
            keep.append(_drop_excluded(x, skus))
        return set(keep) if isinstance(obj, set) else keep
    return obj


def _build_filtered(build, **kw):
    """build_report с отфильтрованными входами. Скаляры (даты, папка, числа)
    проходят как есть — фильтр их не трогает."""
    return build(**{k: _drop_excluded(v, EXCLUDED_SKUS) for k, v in kw.items()})


def _get_netting_cached(run_date, date_from, date_to, recent_days=7):
    """Сверка-выплаты (united-netting) с кэшем на день — чтобы не тянуть тяжёлый
    отчёт каждый прогон. Кэш в netting_cache.json, валиден на текущую дату.
    Попутно (без доп. запроса) строится агрегат за последние `recent_days` дней —
    для колонки «Чистая прибыль ФАКТ ₽/шт за 7д» (в моменте, без сглаживания).
    Returns: (data, data_recent) — оба {(sku,model): {...}}."""
    import json, os
    path = os.path.join(DATA_DIR, "netting_cache.json")
    today = run_date.isoformat()
    stale = None   # прошлый кэш — фолбэк, если перезапрос упадёт
    cached_on = None
    if os.path.exists(path):
        try:
            c = json.load(open(path, encoding="utf-8"))
            if "data_recent" in c:
                stale = ({tuple(k.split("||")): v for k, v in c["data"].items()},
                         {tuple(k.split("||")): v for k, v in c["data_recent"].items()})
            if (c.get("date") == today and c.get("to") == date_to.isoformat()
                    and c.get("recent_days") == recent_days and stale):
                DATA_HEALTH["netting"] = "свежие"
                return stale
            cached_on = c.get("date")
        except Exception:
            pass
    try:
        data, data_recent = fetch_netting(date_from, date_to, recent_days=recent_days,
                                          strict_window=True)
    except Exception as e:
        # Не роняем весь лист сверки — берём прошлый кэш, если есть.
        if stale:
            age = (run_date - date.fromisoformat(cached_on)).days if cached_on else "?"
            print(f"   отчёт по платежам недоступен ({e}); берём прошлый кэш ({len(stale[0])} пар)")
            DATA_HEALTH["netting"] = f"⚠️ из кэша от {cached_on} ({age} дн. назад) — свежий не загрузился"
            return stale
        print(f"   отчёт по платежам недоступен ({e}); сверка будет пустой")
        DATA_HEALTH["netting"] = "❌ НЕТ ДАННЫХ (отчёт по платежам недоступен)"
        return {}, {}
    DATA_HEALTH["netting"] = "свежие"
    try:
        json.dump({"date": today, "from": date_from.isoformat(), "to": date_to.isoformat(),
                   "recent_days": recent_days,
                   "data": {f"{k[0]}||{k[1]}": v for k, v in data.items()},
                   "data_recent": {f"{k[0]}||{k[1]}": v for k, v in data_recent.items()}},
                  open(path, "w", encoding="utf-8"), ensure_ascii=False)
    except Exception:
        pass
    return data, data_recent


def _get_netting_quarter_cached(run_date, date_from, date_to, valid_days=7):
    """Квартальный united-netting с кэшем на 7 дней (netting_quarter_cache.json).
    Квартал меняется медленно → не тянем тяжёлый 90-дневный отчёт каждый день.
    Returns: {(sku,model): {...}}."""
    import json, os
    from datetime import date as _date
    path = os.path.join(DATA_DIR, "netting_quarter_cache.json")
    stale = None   # прошлый кэш — фолбэк, если перезапрос упадёт
    if os.path.exists(path):
        try:
            c = json.load(open(path, encoding="utf-8"))
            stale = {tuple(k.split("||")): v for k, v in c.get("data", {}).items()}
            cached_on = _date.fromisoformat(c.get("date", "2000-01-01"))
            # Валидность — ТОЛЬКО по возрасту (не привязываемся к точному `to`:
            # `to`=вчера меняется каждый день и раньше сбрасывал кэш ежедневно).
            # Квартальная маржа стабильна → освежаем раз в valid_days.
            if (run_date - cached_on).days < valid_days and stale:
                DATA_HEALTH["netting90"] = (f"из кэша от {cached_on.isoformat()} "
                                            f"({(run_date - cached_on).days} дн., обновляется раз в {valid_days})")
                return stale
        except Exception:
            pass
    try:
        data = fetch_netting(date_from, date_to, strict_window=True)
    except Exception as e:
        # Не роняем строки «Σ 90д» — берём прошлый кэш, если он есть.
        if stale:
            print(f"   квартальный нетинг недоступен ({e}); берём прошлый кэш ({len(stale)} пар)")
            DATA_HEALTH["netting90"] = "⚠️ из просроченного кэша — свежий не загрузился"
            return stale
        print(f"   квартальный нетинг недоступен ({e}); квартальные строки будут пустыми")
        DATA_HEALTH["netting90"] = "❌ НЕТ ДАННЫХ"
        return {}
    DATA_HEALTH["netting90"] = "свежие"
    try:
        json.dump({"date": run_date.isoformat(), "from": date_from.isoformat(),
                   "to": date_to.isoformat(),
                   "data": {f"{k[0]}||{k[1]}": v for k, v in data.items()}},
                  open(path, "w", encoding="utf-8"), ensure_ascii=False)
    except Exception:
        pass
    return data


def _record_daily_promos(run_date, promo_prices: dict, write: bool = True):
    """Пишет дневной снимок участия в акциях в promo_daily.json и возвращает всю
    накопленную историю. У Яндекса истории участия в API НЕТ (отдаёт только текущее
    состояние) — накапливаем сами, как ставки буста и остатки FBY. БЕЗ обрезки.

    Структура файла:
      {"meta": {promo_id: {"name","from","to"}},
       "days": {"YYYY-MM-DD": {sku: [promo_id, ...]}}}
    Значение — СПИСОК id, не один: артикул часто участвует в НЕСКОЛЬКИХ акциях
    одновременно (проверено на живых данных: 14 из 26 активных SKU сразу в двух).
    Раньше писали только эффективную (самую дешёвую) акцию — вторая терялась.
    meta храним рядом с днями, чтобы названия/периоды акций пережили их завершение
    (по завершённой акции API уже ничего не отдаст, а в истории она останется).
    Returns: (history_days, meta)."""
    import json, os
    path = os.path.join(DATA_DIR, "promo_daily.json")
    store = {"meta": {}, "days": {}}
    if os.path.exists(path):
        try:
            loaded = json.load(open(path, encoding="utf-8"))
            if isinstance(loaded.get("days"), dict):
                # миграция старого формата (значение — строка id, не список)
                days = {dt: ({sku: ([pid] if isinstance(pid, str) else pid)
                             for sku, pid in day.items()})
                       for dt, day in loaded["days"].items()}
                store = {"meta": loaded.get("meta") or {}, "days": days}
        except Exception:
            pass

    # Акции прочитались не полностью — сегодняшний снимок НЕ пишем: иначе артикулы
    # из непрочитанной акции навсегда останутся в истории как «в этот день не участвовал».
    if not write:
        return store["days"], store["meta"]

    today = run_date.isoformat()
    store["days"][today] = {
        sku: sorted({p["promo_id"] for p in d.get("all_promos", []) if p.get("promo_id")}
                    or ([d["promo_id"]] if d.get("promo_id") else []))
        for sku, d in promo_prices.items()
    }
    store["days"][today] = {sku: pids for sku, pids in store["days"][today].items() if pids}
    for d in promo_prices.values():
        for p in (d.get("all_promos") or [d]):
            pid = p.get("promo_id")
            if pid:
                store["meta"][pid] = {"name": p.get("promo_name", ""),
                                      "from": p.get("period_from", ""),
                                      "to": p.get("period_to", "")}
    try:
        json.dump(store, open(path, "w", encoding="utf-8"), ensure_ascii=False)
    except Exception:
        pass
    return store["days"], store["meta"]


MONTHS_START = date(2026, 5, 1)  # первый месяц с надёжными данными netting


def _get_months_netting_cached(run_date):
    """Помесячный united-netting (сырой агрегат, БЕЗ расчёта прибыли — прибыль
    считается в report_generator, как для Σ 90д). Кэш в netting_months_cache.json:
    завершённые месяцы кэшируются НАВСЕГДА (не меняются задним числом), текущий
    (неполный) месяц тянется заново каждый прогон, пока не закроется.
    Returns: {"YYYY-MM": {(sku,model): {payout,buyer_pay,returns,services,promo,points,other,qty,ret_qty}}}."""
    import json, os, calendar
    path = os.path.join(DATA_DIR, "netting_months_cache.json")
    cache = {}
    if os.path.exists(path):
        try:
            cache = json.load(open(path, encoding="utf-8"))
        except Exception:
            cache = {}

    months = []
    m = MONTHS_START
    cur_month_start = run_date.replace(day=1)
    while m <= cur_month_start:
        months.append(m)
        y, mo = (m.year, m.month + 1) if m.month < 12 else (m.year + 1, 1)
        m = date(y, mo, 1)

    result = {}
    for month_start in months:
        key = month_start.strftime("%Y-%m")
        is_current = (month_start == cur_month_start)
        if not is_current and key in cache:
            result[key] = {tuple(k.split("||")): v for k, v in cache[key].items()}
            continue
        last_day = calendar.monthrange(month_start.year, month_start.month)[1]
        month_end = date(month_start.year, month_start.month, last_day)
        fetch_to = min(month_end, run_date - timedelta(days=1))
        if fetch_to < month_start:
            continue  # месяц ещё не начался
        # Ключ, под которым лежит запасной снимок: у закрытого месяца это он сам,
        # у текущего — отдельный «__partial». Разделение принципиально: неполный
        # снимок текущего месяца, попав под обычный ключ, был бы 1-го числа принят
        # за готовый закрытый месяц и застрял бы там навсегда.
        fb_key = key if not is_current else f"{key}__partial"
        try:
            data = fetch_netting(month_start, fetch_to, strict_window=True)
        except Exception as e:
            print(f"   месяц {key} недоступен ({e}); берём прошлый снимок"
                  if fb_key in cache else f"   месяц {key} недоступен ({e}); пропускаем")
            if fb_key in cache:
                result[key] = {tuple(k.split("||")): v for k, v in cache[fb_key].items()}
            continue
        result[key] = data
        cache[fb_key] = {f"{k[0]}||{k[1]}": v for k, v in data.items()}
        if not is_current:
            cache.pop(f"{key}__partial", None)   # месяц закрылся — черновик не нужен

    try:
        json.dump(cache, open(path, "w", encoding="utf-8"), ensure_ascii=False)
    except Exception:
        pass
    return result


def _record_fby_stock(run_date, stocks):
    """Пишет дневной снимок остатков FBY в fby_stock_daily.json и возвращает всю
    накопленную историю. У Яндекса истории остатков в API НЕТ — накапливаем сами.
    Хранит по SKU: fit (физический) и available (доступный). Один снимок на дату
    (повторный прогон за день перезаписывает). БЕЗ обрезки — храним всё.
    Returns: {"YYYY-MM-DD": {sku: {"fit": int, "available": int}}}."""
    import json, os
    path = os.path.join(DATA_DIR, "fby_stock_daily.json")
    hist = {}
    if os.path.exists(path):
        try:
            hist = json.load(open(path, encoding="utf-8"))
        except Exception:
            hist = {}
    fby = stocks.get("fby", {})
    today = run_date.isoformat()
    hist[today] = {
        sku: {"fit": int(d.get("fit", 0)), "available": int(d.get("available", 0))}
        for sku, d in fby.items()
        if (d.get("fit", 0) or d.get("available", 0))  # не храним пустые
    }
    try:
        json.dump(hist, open(path, "w", encoding="utf-8"), ensure_ascii=False)
    except Exception:
        pass
    return hist


def _get_ums_cached(run_date, date_from, date_to):
    """UMS-отчёт (FEE + буст) с кэшем на день (ums_cache.pkl) — как netting.
    Вечерний прогон не тянет тяжёлый отчёт повторно (лимит 1 генерация/10 мин).
    Pickle, а не JSON — в fee_detail даты живут как datetime, важно сохранить типы.
    Returns: (fee_detail, boost_detail)."""
    import os, pickle
    path = os.path.join(DATA_DIR, "ums_cache.pkl")
    key = (run_date.isoformat(), date_from.isoformat(), date_to.isoformat())
    if os.path.exists(path):
        try:
            with open(path, "rb") as f:
                c = pickle.load(f)
            if c.get("key") == key:
                return c["fee_detail"], c["boost_detail"]
        except Exception:
            pass
    fee_detail, boost_detail = fetch_placement_fee_detail(date_from, date_to)
    try:
        with open(path, "wb") as f:
            pickle.dump({"key": key, "fee_detail": fee_detail,
                         "boost_detail": boost_detail}, f)
    except Exception:
        pass
    return fee_detail, boost_detail


def _df_from_shows_cache(date_from, date_to):
    """Фолбэк: строит df показов из shows_cache.json, когда отчёт shows-sales
    недоступен (3×420 у Яндекса и т.п.). Данные могут быть на день старее,
    но отчёт выйдет, а не упадёт весь прогон."""
    from .shows_cache import load_cache
    rows = []
    for ds, skus in load_cache().items():
        try:
            d = date.fromisoformat(ds)
        except ValueError:
            continue
        if date_from <= d <= date_to:
            for sku, m in skus.items():
                row = {"date": d, "shopSku": sku}
                row.update({k: v for k, v in m.items() if v is not None})
                rows.append(row)
    if not rows:
        return pd.DataFrame(columns=["date", "shopSku", "shows"])
    return pd.DataFrame(rows)


def run(run_date: date):
    print(f"\n=== Yandex Market Analytics Agent — {run_date} ===\n")

    yesterday    = run_date - timedelta(days=1)
    day_before   = run_date - timedelta(days=2)
    week_ago     = run_date - timedelta(days=8)  # 7 дней до вчера
    fourteen_ago = run_date - timedelta(days=14) # 14 дней назад

    # 1. Каталог
    print("1. Загружаем список товаров из каталога...")
    offers = get_all_offers()
    print(f"   Найдено: {len(offers)} товаров")

    # 2. Витрина
    print("2. Парсим витрину...")
    storefront_skus = get_storefront_skus()
    print(f"   На витрине: {len(storefront_skus)} SKU")

    # 3. Отчёт shows-sales за 14 дней (fourteen_ago → yesterday).
    #    best-effort: при недоступности (3×420 и т.п.) строим из кэша показов,
    #    чтобы отчёт вышел, а не падал весь прогон.
    print(f"3. Запрашиваем аналитику за {fourteen_ago} → {yesterday}...")
    try:
        df = fetch_shows_sales(fourteen_ago, yesterday)
        # Normalize date column
        if "День" in df.columns:
            df = df.rename(columns={"День": "date"})
        df["date"] = pd.to_datetime(df["date"], dayfirst=True).dt.date
        DATA_HEALTH["shows"] = "свежие"
    except Exception as e:
        print(f"   ⚠️  shows-sales недоступен ({e}) — строим показы из shows_cache.json")
        df = _df_from_shows_cache(fourteen_ago, yesterday)
        print(f"   Из кэша: {len(df)} строк")
        DATA_HEALTH["shows"] = "⚠️ ИЗ КЭША — shows-sales не загрузился, воронка может быть неполной"

    # Кэш показов по дням: чиним нулевые «дыры» Яндекса (день был, в снимке стал 0).
    # Ненулевые дни перезаписывают кэш; нулевые — берутся из кэша.
    before_zeros = int((df.get("shows", pd.Series(dtype=float)).fillna(0) == 0).sum()) \
        if "shows" in df.columns else 0
    df = merge_with_cache(df, fourteen_ago, yesterday)
    after_zeros = int((df.get("shows", pd.Series(dtype=float)).fillna(0) == 0).sum()) \
        if "shows" in df.columns else 0
    print(f"   Кэш показов: было нулевых строк {before_zeros}, после подстановки {after_zeros}")

    def day_df(d: date) -> pd.DataFrame:
        return df[df["date"] == d].copy()

    df_yesterday  = day_df(yesterday)
    df_day_before = day_df(day_before)
    df_week_ago   = day_df(week_ago)

    # 7-дневная скользящая (все дни кроме вчера)
    df_rolling = df[df["date"] < yesterday].copy()

    def rolling_avg(sku: str, metric: str) -> float:
        if df_rolling.empty or "shopSku" not in df_rolling.columns or metric not in df_rolling.columns:
            return 0.0
        sub = df_rolling[df_rolling["shopSku"] == sku][metric].dropna()
        return float(sub.mean()) if len(sub) else 0.0

    # 4. Фильтрация: только товары с показами хотя бы в одном из трёх дней
    def has_shows(sku, *dfs):
        for d in dfs:
            if d.empty or "shopSku" not in d.columns:
                continue
            row = d[d["shopSku"] == sku]
            if not row.empty and "shows" in row.columns:
                val = row["shows"].fillna(0).sum()
                if val > 0:
                    return True
        return False

    active_offers = [o for o in offers if has_shows(o["shopSku"], df_yesterday, df_day_before, df_week_ago)]
    print(f"   Активных товаров (с показами): {len(active_offers)} из {len(offers)}")

    # 5. Алерты
    print("4. Анализируем показатели...")
    alerts = []

    y_idx = df_yesterday.set_index("shopSku").to_dict("index") if not df_yesterday.empty else {}
    db_idx = df_day_before.set_index("shopSku").to_dict("index") if not df_day_before.empty else {}

    for offer in active_offers:
        sku  = offer.get("shopSku", "")
        name = offer.get("name", sku)
        y    = y_idx.get(sku, {})
        db   = db_idx.get(sku, {})

        shows_y  = (y.get("shows", 0) or 0)
        shows_db = (db.get("shows", 0) or 0)
        avg      = rolling_avg(sku, "shows")
        on_sf    = sku in storefront_skus

        # Нет на витрине но вчера были показы
        if shows_y > 0 and not on_sf and storefront_skus:
            alerts.append({
                "type": "absent",
                "sku": sku,
                "name": name,
                "msg": f"Вчера {int(shows_y)} показов, но товара нет на витрине",
            })

        # На витрине но 0 показов
        if on_sf and shows_y == 0:
            alerts.append({
                "type": "zero",
                "sku": sku,
                "name": name,
                "msg": "0 показов вчера, но товар на витрине!",
            })

        # Падение показов: вчера относительно позавчера И только если падение
        # выходит за рамки обычных дневных колебаний (> 1 σ дневного ряда).
        if shows_db > 0 and shows_y < shows_db:
            drop_abs = shows_db - shows_y
            pct_drop = round((1 - shows_y / shows_db) * 100)
            # дневной ряд показов по этому SKU (для оценки нормальных колебаний)
            series = []
            if not df.empty and "shopSku" in df.columns and "shows" in df.columns:
                series = df[df["shopSku"] == sku]["shows"].fillna(0).tolist()
            std = statistics.pstdev(series) if len(series) >= 3 else 0
            beyond_noise = (drop_abs > std) if std > 0 else (pct_drop >= 50)
            if beyond_noise and pct_drop >= 30:
                alerts.append({
                    "type": "drop",
                    "sku": sku,
                    "name": name,
                    "msg": (f"{int(shows_y)} показов вчера vs {int(shows_db)} позавчера "
                            f"(−{pct_drop}%, за пределами обычных колебаний)"),
                })

    # «На витрине, но показов нет вообще» — цикл выше этот случай пропускал:
    # active_offers отбирает только товары с показами хотя бы в один из трёх
    # контрольных дней, поэтому товар, молчащий всю неделю (самый тревожный
    # случай), в алерты не попадал в принципе.
    # Условие на непустые данные обязательно: при недоступном shows-sales агент
    # откатывается на тонкий кэш, и без него алерт сработал бы по всему каталогу.
    y_shows = df_yesterday["shows"].fillna(0).sum() \
        if not df_yesterday.empty and "shows" in df_yesterday.columns else 0
    if storefront_skus and y_shows > 0:
        active_skus = {o.get("shopSku") for o in active_offers}
        for offer in offers:
            sku = offer.get("shopSku", "")
            if sku in storefront_skus and sku not in active_skus:
                alerts.append({
                    "type": "zero",
                    "sku": sku,
                    "name": offer.get("name", sku),
                    "msg": "0 показов вчера, позавчера и неделю назад, но товар на витрине!",
                })

    if alerts:
        print(f"\n   ⚠️  Алертов: {len(alerts)}")
        for a in alerts:
            print(f"      {'🔴' if a['type'] in ('zero','absent') else '🟡'} {a['name']}: {a['msg']}")
    else:
        print("   ✅ Все показатели в норме")

    # 5. FBS/FBY заказы и остатки
    print("\n5. Загружаем заказы FBS/FBY и остатки...")
    orders_by_type = get_orders_by_type(week_ago, yesterday)
    print(f"   FBS дней с заказами: {len(orders_by_type['fbs'])}, FBY: {len(orders_by_type['fby'])}")

    stocks = get_stocks()
    print(f"   FBS SKU с остатками: {len(stocks['fbs'])}, FBY: {len(stocks['fby'])}")

    # Дневной снимок FBY → накапливаем историю остатков (API истории не даёт)
    fby_stock_history = _record_fby_stock(run_date, stocks)
    print(f"   История FBY: {len(fby_stock_history)} дней в fby_stock_daily.json")

    # FBY restock estimation (7-day window)
    fby_restock = estimate_fby_restock(stocks["fby"], orders_by_type["fby"], num_days=7)
    urgent = [r for r in fby_restock if r["urgency"].startswith("🔴")]
    if urgent:
        print(f"   🔴 Срочная поставка нужна для {len(urgent)} SKU:")
        for r in urgent:
            print(f"      {r['sku']}: осталось {r['available']} шт, "
                  f"ср. продаж {r['avg_daily']}/день, хватит на {r['days_left']} дн.")
        # В алерты дашборда поставку НЕ добавляем (по запросу) — остаётся
        # на отдельной вкладке остатков/поставок.

    # 6. Юнит-экономика: заказы с комиссиями + каталожные цены
    #    Для логистики по регионам берём 20 дней
    twenty_ago = run_date - timedelta(days=20)
    print("\n6. Загружаем заказы с комиссиями (юнит-экономика, 20 дней)...")
    all_fee_orders = get_orders_with_fees(twenty_ago, yesterday)
    fbs_cnt = len(all_fee_orders["fbs"])
    fby_cnt = len(all_fee_orders["fby"])
    print(f"   Позиций с комиссиями: FBS={fbs_cnt}, FBY={fby_cnt}")

    print("   Загружаем каталожные цены...")
    catalog_prices = get_catalog_prices()
    print(f"   Цены для {len(catalog_prices)} SKU")

    print("   Загружаем акционные цены...")
    # Best-effort: сбой API акций не должен ронять весь отчёт (16–17.09.2026 два
    # прогона подряд упали на 500 от /promos/offers по новой акции Яндекса).
    promo_ok = True
    try:
        promo_prices = get_promo_prices()
        DATA_HEALTH["promo"] = "свежие"
    except PromoFetchPartial as e:
        promo_prices, promo_ok = e.result, False
        DATA_HEALTH["promo"] = f"⚠️ неполные — не загрузились: {', '.join(e.failed)}"
    except Exception as e:
        promo_prices, promo_ok = {}, False
        print(f"   ⚠️  акции недоступны ({str(e)[:120]})")
        DATA_HEALTH["promo"] = "❌ НЕТ ДАННЫХ (API акций недоступен)"
    print(f"   В акциях: {len(promo_prices)} SKU")

    # Дневной снимок участия в акциях → накапливаем историю (API её не даёт)
    promo_history, promo_meta = _record_daily_promos(run_date, promo_prices, write=promo_ok)
    if not promo_ok:
        print("   Снимок участия за сегодня НЕ записан — акции прочитаны не полностью")
    print(f"   История акций: {len(promo_history)} дней в promo_daily.json "
          f"({len(promo_meta)} акций в справочнике)")

    # UMS-отчёт: детализация FEE + РЕАЛЬНЫЙ буст по SKU из листа «Буст продаж»
    # (сработавшая ставка обычно сильно ниже потолка). Окно 14 дней — компромисс
    # между свежестью (чтобы видеть эффект изменения ставок) и покрытием
    # (хватает продаж, чтобы не было сплошных нулей).
    ums_from = run_date - timedelta(days=14)
    print(f"   Загружаем UMS (FEE + буст) за {ums_from} → {yesterday} (кэш на день)...")
    try:
        fee_detail, boost_detail = _get_ums_cached(run_date, ums_from, yesterday)
    except Exception as e:
        print(f"   UMS недоступен ({e}); FEE-детализация и факт-буст будут пустыми")
        fee_detail, boost_detail = [], {}
    print(f"   Заказов в детализации FEE: {len(fee_detail)}; SKU с фактич. бустом: {len(boost_detail)}")

    # 6b. Новые срезы: индекс цен + отзывы (качество карточек уже в offers)
    print("\n6b. Загружаем привлекательность цен...")
    price_reco = get_price_recommendations()
    print(f"   Рекомендации цен для {len(price_reco)} SKU")

    print("   Загружаем отзывы...")
    try:
        reviews = get_reviews_summary()
    except Exception as e:
        print(f"   отзывы недоступны ({e}); пропускаем (нужен scope «Общение с покупателями»)")
        reviews = {}
    print(f"   SKU с отзывами: {len(reviews)}")

    # 6c. Буст: снимок текущих ставок из кабинета (boost_bids_history.csv,
    #     обновляется вручную) + дневная история API-ставок.
    #     Рекомендации ставок (get_bid_recommendations) больше НЕ тянем: их
    #     потреблял только лист «Ставки буста», удалён 30.07.2026 как
    #     неинформативный. Лист «Ставки буста ПО ДНЯМ» — другой, он живёт.
    print("\n6c. Загружаем ставки буста (снимок + история)...")
    active_skus = [o["shopSku"] for o in active_offers]
    boost_bids, boost_bids_date = _load_boost_snapshot()
    print(f"   Снимок текущих ставок: {len(boost_bids)} SKU (от {boost_bids_date or '—'})")
    # Авто-снимок ставок через API (bids/info) в дневную историю → «ставка вчера».
    # Пока буст ведётся кампаниями кабинета, bids/info пуст; заполнится после
    # перевода на API (пилот).
    try:
        api_bids_today = get_current_bids(active_skus)
    except Exception as e:
        print(f"   bids/info недоступен ({e})")
        api_bids_today = {}
    boost_bids_prev, boost_bids_history = _record_daily_bids(run_date, api_bids_today)
    print(f"   API-ставок сегодня: {len(api_bids_today)}; вчера было: {len(boost_bids_prev)}")

    # 6d. Сверка с Яндексом — реальные выплаты по артикулам (united-netting),
    #     trailing 30 дней, кэш на день. Источник истины «что Яндекс перевёл».
    netting_from = run_date - timedelta(days=30)
    print(f"\n6d. Сверка выплат (united-netting) за {netting_from} → {yesterday}...")
    netting, netting_recent = _get_netting_cached(run_date, netting_from, yesterday, recent_days=7)
    print(f"   Пар (SKU, модель) в выплатах: {len(netting)} (за 7д: {len(netting_recent)})")

    # 6e. Квартальная сверка (90 дней) — устойчивая маржа без дневных скачков от
    #     смещения выплат. Кэш на 7 дней (квартал меняется медленно, не тянем ежедневно).
    netting_q_from = run_date - timedelta(days=90)
    print(f"6e. Квартальная сверка (90д) за {netting_q_from} → {yesterday}...")
    netting_quarter = _get_netting_quarter_cached(run_date, netting_q_from, yesterday)
    print(f"   Пар (SKU, модель) за квартал: {len(netting_quarter)}")

    # 6f. Помесячная сверка — сырой агрегат по месяцам (завершённые месяцы кэшируются
    # навсегда), профит считает report_generator (та же формула, что и Σ90д).
    print("6f. Помесячная сверка (netting по месяцам)...")
    netting_months = _get_months_netting_cached(run_date)
    print(f"   Месяцев в разбивке: {len(netting_months)} ({', '.join(sorted(netting_months))})")

    # API-ставки (если заданы per-SKU) важнее ручного снимка кабинета.
    for sku, bid in api_bids_today.items():
        e = boost_bids.setdefault(sku, {})
        e["bid_pct"] = bid
        e["campaign"] = "Кампания (API)"

    # 7. Excel
    print("\n7. Генерируем Excel-отчёт...")
    if EXCLUDED_SKUS:
        print(f"   Скрываем снятые с продажи: {', '.join(sorted(EXCLUDED_SKUS))}")
    report_path = _build_filtered(build_report,
        run_date=run_date,
        active_offers=active_offers,
        all_offers=offers,
        df_full=df,
        df_yesterday=df_yesterday,
        df_day_before=df_day_before,
        df_week_ago=df_week_ago,
        df_rolling=df_rolling,
        storefront_skus=storefront_skus,
        alerts=alerts,
        orders_by_type=orders_by_type,
        stocks=stocks,
        fby_stock_history=fby_stock_history,
        fby_restock=fby_restock,
        report_folder=REPORT_FOLDER,
        data_health=DATA_HEALTH,
        all_fee_orders=all_fee_orders,
        catalog_prices=catalog_prices,
        promo_prices=promo_prices,
        fee_detail=fee_detail,
        boost_detail=boost_detail,
        fee_window_days=(yesterday - twenty_ago).days,
        date_from=fourteen_ago,
        date_to=yesterday,
        price_reco=price_reco,
        reviews=reviews,
        boost_bids=boost_bids,
        boost_bids_date=boost_bids_date,
        boost_bids_prev=boost_bids_prev,
        boost_bids_history=boost_bids_history,
        netting=netting,
        netting_from=netting_from,
        netting_to=yesterday,
        netting_recent=netting_recent,
        netting_recent_days=7,
        netting_quarter=netting_quarter,
        netting_quarter_from=netting_q_from,
        netting_quarter_to=yesterday,
        netting_months=netting_months,
        promo_history=promo_history,
        promo_meta=promo_meta,
    )
    print(f"   Отчёт: {report_path}")
    print("\n=== Готово ===\n")
    return report_path, alerts


if __name__ == "__main__":
    _enable_timestamps()
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", help="YYYY-MM-DD, default=today")
    args = parser.parse_args()
    run_date = date.fromisoformat(args.date) if args.date else date.today()

    try:
        run(run_date)
    except Exception:
        print("\n[ERROR]")
        traceback.print_exc()
        sys.exit(1)
