"""
Excel report. Основные листы:
  Дашборд            — сводка + наглядные алерты
  Товары             — per-SKU воронка: вчера / позавчера / нед.назад + динамика
  Индекс цен         — привлекательность цен (выгодная/умеренная/завышенная) + рекомендации
  Дырявая воронка    — SKU с высоким трафиком и низкой конверсией корзина→заказ
  Качество и отзывы  — контент-рейтинг карточек + агрегированные отзывы/рейтинг
  Юнит-экономика, Скидки FEE, Видимость↔Буст↔Показы, Акции vs Без акций, Остатки и др.
"""

import os
import statistics
from datetime import date, timedelta

import numpy as np
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.comments import Comment


# ── Palette ───────────────────────────────────────────────────────────────────
RED_BG     = "FF4444"
RED_LIGHT  = "FFC7CE"
ORANGE_BG  = "FF8C00"
ORANGE_LT  = "FFE0B2"
YELLOW_LT  = "FFF9C4"
GREEN_LT   = "C6EFCE"
GREEN_BG   = "27AE60"
BLUE_BG    = "2F4F8F"
BLUE_LT    = "BDD7EE"
GREY_LT    = "F2F2F2"
WHITE      = "FFFFFF"

def load_ozon_ppu(report_folder: str) -> tuple[dict, str]:
    """Прибыль ₽/ед по артикулам из САМОГО СВЕЖЕГО отчёта Озона в той же папке.

    Читаем готовый xlsx, а не считаем заново: у Озона своя экономика (комиссия
    ~45–50%, свои услуги, свой источник данных — выгрузка начислений), повторять
    её здесь нельзя. Берём то, что уже посчитал ozon-агент.
    Returns: ({offer_id: ₽/ед}, дата отчёта 'YYYY-MM-DD' или '').
    """
    import glob
    import openpyxl
    files = sorted(glob.glob(os.path.join(report_folder, "ozon_report_*.xlsx")))
    if not files:
        return {}, ""
    path = files[-1]
    rep_date = os.path.basename(path).replace("ozon_report_", "").replace(".xlsx", "")
    try:
        ws = openpyxl.load_workbook(path, data_only=True, read_only=True)["Юнит-экономика"]
    except Exception:
        return {}, ""
    rows = list(ws.iter_rows(min_row=3, values_only=True))
    if not rows:
        return {}, ""
    head = rows[0]
    try:
        oid_col = [i for i, h in enumerate(head) if str(h).strip() == "offer_id"][0]
        ppu_col = [i for i, h in enumerate(head)
                   if h and "Прибыль" in str(h) and "/ед" in str(h)][0]
    except IndexError:
        return {}, ""
    out = {}
    for row in rows[1:]:
        if len(row) <= max(oid_col, ppu_col):
            continue
        oid, val = row[oid_col], row[ppu_col]
        if oid and isinstance(val, (int, float)):
            out[str(oid)] = round(float(val), 2)
    return out, rep_date


def _fill(c):    return PatternFill("solid", fgColor=c)

def _blend_hex(c_from, c_to, t):
    """Линейная интерполяция между двумя HEX-цветами, t∈[0,1]."""
    t = max(0.0, min(1.0, t))
    f = tuple(int(c_from[i:i+2], 16) for i in (0, 2, 4))
    to = tuple(int(c_to[i:i+2], 16) for i in (0, 2, 4))
    mixed = tuple(round(f[i] + (to[i] - f[i]) * t) for i in range(3))
    return "".join(f"{v:02X}" for v in mixed)

def _gradient_fill(pct_change, cap=100.0):
    """Заливка-градиент по % изменению: рост → к зелёному, падение → к красному,
    интенсивность растёт с величиной изменения (капается на `cap`%). None → без заливки."""
    if pct_change is None:
        return None
    if pct_change >= 0:
        t = min(pct_change / cap, 1.0)
        return _fill(_blend_hex("FFFFFF", GREEN_BG, t))
    t = min(abs(pct_change) / cap, 1.0)
    return _fill(_blend_hex("FFFFFF", "C0392B", t))
def _hfont():    return Font(bold=True, color="FFFFFF", size=11)
def _tfont(sz=11, bold=False, color="000000"): return Font(bold=bold, color=color, size=sz)
def _border():
    s = Side(style="thin", color="CCCCCC")
    return Border(left=s, right=s, top=s, bottom=s)
def _center():   return Alignment(horizontal="center", vertical="center", wrap_text=True)
def _left():     return Alignment(horizontal="left", vertical="center", wrap_text=True)

def _pct(a, b):
    try: return round(a / b * 100, 1) if b else None
    except: return None

def _delta(today, ref):
    try:
        if not ref or ref == 0: return "—"
        d = round((today - ref) / ref * 100, 1)
        return f"+{d}%" if d >= 0 else f"{d}%"
    except: return "—"

def _v(d: dict, key: str, default=0):
    val = d.get(key, default)
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return default
    try: return float(val)
    except: return default


def _vnone(d: dict, key: str):
    """Like _v but returns None for missing/NaN (distinguishes 'no data' from 0)."""
    val = d.get(key)
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try: return float(val)
    except (TypeError, ValueError): return None


def _style_header_row(ws, row_num, num_cols, bg=BLUE_BG):
    for col in range(1, num_cols + 1):
        cell = ws.cell(row=row_num, column=col)
        cell.font = _hfont()
        cell.fill = _fill(bg)
        cell.alignment = _center()
        cell.border = _border()


def build_report(
    run_date: date,
    active_offers: list[dict],
    all_offers: list[dict],
    df_full: pd.DataFrame,
    df_yesterday: pd.DataFrame,
    df_day_before: pd.DataFrame,
    df_week_ago: pd.DataFrame,
    df_rolling: pd.DataFrame,
    storefront_skus: set[str],
    alerts: list[dict],
    orders_by_type: dict,
    stocks: dict,
    fby_restock: list[dict],
    report_folder: str,
    **kwargs,
) -> str:
    os.makedirs(report_folder, exist_ok=True)

    yesterday  = run_date - timedelta(days=1)
    day_before = run_date - timedelta(days=2)
    week_ago   = run_date - timedelta(days=8)

    # Index by shopSku
    def idx(df):
        return df.set_index("shopSku").to_dict("index") if not df.empty and "shopSku" in df.columns else {}
    y_idx  = idx(df_yesterday)
    db_idx = idx(df_day_before)
    w_idx  = idx(df_week_ago)

    def rolling_avg(sku, metric):
        if df_rolling.empty or "shopSku" not in df_rolling.columns or metric not in df_rolling.columns:
            return 0.0
        sub = df_rolling[df_rolling["shopSku"] == sku][metric].dropna()
        return float(sub.mean()) if len(sub) else 0.0

    wb = Workbook()

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 1: ДАШБОРД
    # ══════════════════════════════════════════════════════════════════════════
    ws1 = wb.active
    ws1.title = "Дашборд"
    ws1.sheet_properties.tabColor = "2F4F8F"

    # Title
    ws1.merge_cells("A1:H1")
    ws1["A1"] = f"Яндекс Маркет — аналитика продаж · {run_date.strftime('%d.%m.%Y')}"
    ws1["A1"].font = Font(bold=True, color="FFFFFF", size=16)
    ws1["A1"].fill = _fill(BLUE_BG)
    ws1["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws1.row_dimensions[1].height = 40

    # ── Свежесть источников ──
    # Отчёт может собраться успешно и выглядеть нормально, будучи построенным на
    # вчерашних данных (кэш подставляется молча, когда API не ответил). Строка
    # показывает, что реально пошло в файл; жёлтая — если хоть где-то фолбэк.
    health = kwargs.get("data_health") or {}
    if health:
        labels = {"shows": "Воронка", "promo": "Акции", "netting": "Выплаты 30д", "netting90": "Выплаты 90д"}
        parts = [f"{labels.get(k, k)}: {v}" for k, v in health.items() if k in labels]
        degraded = any(("⚠️" in v or "❌" in v) for v in health.values())
        ws1.merge_cells("A2:H2")
        ws1["A2"] = "Данные — " + " · ".join(parts)
        ws1["A2"].font = Font(bold=degraded, size=9,
                              color="B71C1C" if degraded else "666666")
        if degraded:
            ws1["A2"].fill = _fill(YELLOW_LT)
        ws1["A2"].alignment = Alignment(horizontal="left", vertical="center")

    # ── Totals row ──
    ws1.append([])
    row = 3
    headers = ["Период", "Показы", "Клики", "CTR", "В корзину", "CR показ→корз", "Заказы", "Заказы с бустом"]
    for c, h in enumerate(headers, 1):
        ws1.cell(row=row, column=c, value=h)
    _style_header_row(ws1, row, len(headers))

    def total_row(label, df):
        if df.empty:
            return [label, 0, 0, "—", 0, "—", 0, 0]
        shows  = df["shows"].fillna(0).sum() if "shows" in df.columns else 0
        clicks = df["clicks"].fillna(0).sum() if "clicks" in df.columns else 0
        carts  = df["carts"].fillna(0).sum() if "carts" in df.columns else 0
        orders = df["orders"].fillna(0).sum() if "orders" in df.columns else 0
        o_boost= df["orders_promo"].fillna(0).sum() if "orders_promo" in df.columns else 0
        ctr    = _pct(clicks, shows)
        cr     = _pct(carts, shows)
        return [label, int(shows), int(clicks), f"{ctr}%" if ctr else "—",
                int(carts), f"{cr}%" if cr else "—", int(orders), int(o_boost)]

    for label, df in [
        (f"Вчера ({yesterday.strftime('%d.%m')})", df_yesterday),
        (f"Позавчера ({day_before.strftime('%d.%m')})", df_day_before),
        (f"Нед. назад ({week_ago.strftime('%d.%m')})", df_week_ago),
    ]:
        row += 1
        vals = total_row(label, df)
        for c, v in enumerate(vals, 1):
            cell = ws1.cell(row=row, column=c, value=v)
            cell.border = _border()
            cell.alignment = _center() if c > 1 else _left()
            if c == 1:
                cell.font = _tfont(bold=True)

    # ── Динамика ──
    row += 2
    ws1.cell(row=row, column=1, value="Динамика").font = _tfont(sz=13, bold=True)
    row += 1
    dyn_headers = ["", "Показы", "Клики", "В корзину", "Заказы"]
    for c, h in enumerate(dyn_headers, 1):
        ws1.cell(row=row, column=c, value=h)
    _style_header_row(ws1, row, len(dyn_headers), "555555")

    def safe_sum(df, col):
        if df.empty or col not in df.columns: return 0
        return df[col].fillna(0).sum()

    metrics = ["shows", "clicks", "carts", "orders"]
    for label, d_a, d_b in [
        ("Вчера vs позавчера", df_yesterday, df_day_before),
        ("Вчера vs нед. назад", df_yesterday, df_week_ago),
    ]:
        row += 1
        ws1.cell(row=row, column=1, value=label).font = _tfont(bold=True)
        ws1.cell(row=row, column=1).border = _border()
        for i, m in enumerate(metrics):
            a = safe_sum(d_a, m)
            b = safe_sum(d_b, m)
            delta = _delta(a, b)
            cell = ws1.cell(row=row, column=i + 2, value=delta)
            cell.alignment = _center()
            cell.border = _border()
            cell.font = _tfont(bold=True)
            # colour code
            if delta.startswith("+"):
                cell.fill = _fill(GREEN_LT)
                cell.font = _tfont(bold=True, color="1B5E20")
            elif delta.startswith("-"):
                cell.fill = _fill(RED_LIGHT)
                cell.font = _tfont(bold=True, color="B71C1C")

    # ── РОСТ ПОКАЗОВ (вчера vs позавчера, порог > 20% от среднего) ──
    # Средние показы по всем активным позициям за вчера
    avg_shows_all = 0
    if not df_yesterday.empty and "shows" in df_yesterday.columns:
        avg_shows_all = df_yesterday["shows"].fillna(0).mean()

    growth_threshold = avg_shows_all * 0.2  # 20% от среднего

    growth_items = []
    for offer in active_offers:
        sku = offer.get("shopSku", "")
        name = offer.get("name", sku)
        y = y_idx.get(sku, {})
        db = db_idx.get(sku, {})
        shows_y = _v(y, "shows")
        shows_db = _v(db, "shows")
        diff = shows_y - shows_db
        if diff > growth_threshold and shows_db > 0:
            pct = round(diff / shows_db * 100, 1)
            growth_items.append({
                "name": name, "sku": sku,
                "shows_y": int(shows_y), "shows_db": int(shows_db),
                "diff": int(diff), "pct": pct,
            })
    growth_items.sort(key=lambda x: -x["diff"])

    row += 2
    ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    cell = ws1.cell(row=row, column=1,
        value=f"📈  РОСТ ПОКАЗОВ  ({len(growth_items)} позиций, порог +{int(growth_threshold)} шт = 20% от ср.)")
    cell.font = Font(bold=True, color="FFFFFF", size=13)
    cell.fill = _fill(GREEN_BG)
    cell.alignment = Alignment(horizontal="center")
    ws1.row_dimensions[row].height = 32

    if growth_items:
        row += 1
        g_heads = ["", "Товар", "SKU", f"Вчера\n{yesterday.strftime('%d.%m')}",
                   f"Позавчера\n{day_before.strftime('%d.%m')}", "Рост шт.", "Рост %"]
        for c, h in enumerate(g_heads, 1):
            cell = ws1.cell(row=row, column=c, value=h)
            cell.font = _tfont(sz=10, bold=True)
            cell.fill = _fill(GREY_LT)
            cell.border = _border()

        for i, g in enumerate(growth_items, 1):
            row += 1
            vals = [i, g["name"][:45], g["sku"], g["shows_y"], g["shows_db"],
                    f"+{g['diff']}", f"+{g['pct']}%"]
            for c, v in enumerate(vals, 1):
                cell = ws1.cell(row=row, column=c, value=v)
                cell.fill = _fill(GREEN_LT)
                cell.border = _border()
                cell.alignment = _left() if c in (2, 3) else _center()
                if c >= 6:
                    cell.font = _tfont(bold=True, color="1B5E20")
    else:
        row += 1
        ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        cell = ws1.cell(row=row, column=1, value="Нет позиций с ростом > 20% от среднего")
        cell.alignment = Alignment(horizontal="center")
        cell.fill = _fill(GREY_LT)

    # ── АЛЕРТЫ (наглядный блок) ──
    row += 2
    ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
    ws1.cell(row=row, column=1, value="АЛЕРТЫ").font = Font(bold=True, color="FFFFFF", size=14)
    ws1.cell(row=row, column=1).fill = _fill(RED_BG if alerts else GREEN_BG)
    ws1.cell(row=row, column=1).alignment = Alignment(horizontal="center")
    ws1.row_dimensions[row].height = 32

    if not alerts:
        row += 1
        ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
        cell = ws1.cell(row=row, column=1, value="✅  Все показатели в норме — аномалий не обнаружено")
        cell.font = _tfont(sz=12, bold=True, color="1B5E20")
        cell.fill = _fill(GREEN_LT)
        cell.alignment = Alignment(horizontal="center")
    else:
        # Group alerts by type
        groups = {
            "zero":    {"icon": "🔴", "title": "0 показов на витрине", "bg": RED_LIGHT, "hdr": RED_BG, "items": []},
            "absent":  {"icon": "⚠️",  "title": "Исчезли с витрины", "bg": ORANGE_LT, "hdr": ORANGE_BG, "items": []},
            "drop":    {"icon": "🟡", "title": "Падение показов", "bg": YELLOW_LT, "hdr": "F9A825", "items": []},
        }
        for a in alerts:
            t = a.get("type", "drop")
            if t in groups:
                groups[t]["items"].append(a)

        for gtype in ["zero", "absent", "drop"]:
            g = groups[gtype]
            if not g["items"]:
                continue

            row += 1
            ws1.merge_cells(start_row=row, start_column=1, end_row=row, end_column=8)
            cell = ws1.cell(row=row, column=1,
                value=f'{g["icon"]}  {g["title"]}  ({len(g["items"])})')
            cell.font = Font(bold=True, color="FFFFFF", size=12)
            cell.fill = _fill(g["hdr"])
            cell.alignment = Alignment(horizontal="left", indent=1)
            ws1.row_dimensions[row].height = 28

            # Sub-header
            row += 1
            sub_h = ["", "Товар", "SKU", "Вчера показы", "Позавчера", "Ср.7д", "Описание"]
            for c, h in enumerate(sub_h, 1):
                cell = ws1.cell(row=row, column=c, value=h)
                cell.font = _tfont(sz=10, bold=True)
                cell.fill = _fill(GREY_LT)
                cell.border = _border()

            for i, a in enumerate(g["items"], 1):
                row += 1
                y_data = y_idx.get(a["sku"], {})
                db_data = db_idx.get(a["sku"], {})
                vals = [
                    i,
                    a["name"][:45],
                    a["sku"],
                    int(_v(y_data, "shows")),
                    int(_v(db_data, "shows")),
                    round(rolling_avg(a["sku"], "shows"), 0),
                    a["msg"],
                ]
                for c, v in enumerate(vals, 1):
                    cell = ws1.cell(row=row, column=c, value=v)
                    cell.fill = _fill(g["bg"])
                    cell.border = _border()
                    cell.alignment = _left() if c in (2, 3, 7) else _center()

    # Column widths for Sheet 1
    ws1.column_dimensions["A"].width = 28
    ws1.column_dimensions["B"].width = 38
    ws1.column_dimensions["C"].width = 16
    ws1.column_dimensions["D"].width = 16
    ws1.column_dimensions["E"].width = 16
    ws1.column_dimensions["F"].width = 16
    ws1.column_dimensions["G"].width = 18
    ws1.column_dimensions["H"].width = 18

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 2: ТОВАРЫ
    # ══════════════════════════════════════════════════════════════════════════
    ws2 = wb.create_sheet("Товары")
    ws2.sheet_properties.tabColor = "2980B9"

    # Stock lookup: FBS и FBY отдельно (available, по всем складам)
    fbs_stocks_map = stocks.get("fbs", {})
    fby_stocks_map = stocks.get("fby", {})

    heads = [
        "Товар", "SKU", "Остаток\nFBS", "Остаток\nFBY",
        f"Показы\n{yesterday.strftime('%d.%m')}", f"Показы\n{day_before.strftime('%d.%m')}",
        f"Показы\n{week_ago.strftime('%d.%m')}", "Ср.7д\nпоказы",
        "Δ вч.\nvs позавч.", "Δ вч.\nvs нед.",
        "% показов\nс бустом\nвчера", "% показов\nс бустом\nср.7д",
        f"Клики\n{yesterday.strftime('%d.%m')}", f"Клики\n{day_before.strftime('%d.%m')}",
        "CTR%\nвчера", "CTR%\nпозавч.",
        f"Корзина\n{yesterday.strftime('%d.%m')}", f"Корзина\n{day_before.strftime('%d.%m')}",
        "CTR карт.\nвчера", "CTR карт.\nпозавч.",
        "% в корзину\nвчера", "% в корзину\nпозавч.",
        f"Заказы\n{yesterday.strftime('%d.%m')}", f"Заказы\n{day_before.strftime('%d.%m')}",
        "Буст\nвчера", "% буста\nвчера",
        "Статус",
    ]
    # COL indices (1-based): Остаток FBS=3, Остаток FBY=4, Показы вчера=5, Показы позавч=6,
    # Показы нед=7, Ср7д=8, Δ позавч=9, Δ нед=10, %пок.буст вч=11, %пок.буст ср7д=12,
    # Клики вч=13, Клики позавч=14, CTR вч=15, CTR позавч=16,
    # Корзина вч=17, Корзина позавч=18, CTR карт.(дубль)вч=19, CTR карт.позавч=20,
    # %в корзину вч=21, %в корзину позавч=22, Заказы вч=23, Заказы позавч=24,
    # Буст=25, %буста=26, Статус=27
    # (20.09.2026 убраны «На витрине» — дублировала «Статус», где тот же факт виден
    #  как «🔴 Нет на витрине», — и четыре колонки про акции: участие и цены акции
    #  смотрим на «Юнит-экономике», здесь они были не нужны. Сами данные по акциям
    #  агент по-прежнему тянет — они нужны истории participation и другим листам.)
    # НОВЫЕ КОЛОНКИ ДОБАВЛЯТЬ ТОЛЬКО В КОНЕЦ — лист хрупок к сдвигам (COL_* ниже).
    COL_STOCK_FBS, COL_STOCK_FBY = 3, 4
    COL_SHOWS_Y, COL_SHOWS_DB = 5, 6
    COL_DELTA1, COL_DELTA2 = 9, 10
    COL_BSHOW_Y, COL_BSHOW_AVG = 11, 12
    COL_CLICKS_Y, COL_CLICKS_DB = 13, 14
    COL_CARTS_Y, COL_CARTS_DB = 17, 18
    COL_ORDERS_Y, COL_ORDERS_DB = 23, 24

    ws2.append(heads)
    _style_header_row(ws2, 1, len(heads))
    ws2.row_dimensions[1].height = 36

    # Collect data for cross-sectional correlation
    corr_stock = []
    corr_shows = []
    corr_clicks = []
    corr_carts = []
    corr_orders = []

    for offer in active_offers:
        sku  = offer.get("shopSku", "")
        name = offer.get("name", sku)
        y = y_idx.get(sku, {})
        db = db_idx.get(sku, {})
        w = w_idx.get(sku, {})

        shows_y = _v(y, "shows"); shows_db = _v(db, "shows"); shows_w = _v(w, "shows")
        shows_avg = rolling_avg(sku, "shows")
        clicks_y = _v(y, "clicks"); clicks_db = _v(db, "clicks")
        clicks_avg = rolling_avg(sku, "clicks")
        carts_y = _v(y, "carts"); carts_db = _v(db, "carts")
        carts_avg = rolling_avg(sku, "carts")
        orders_y = _v(y, "orders"); orders_db = _v(db, "orders")
        orders_avg = rolling_avg(sku, "orders")
        boost_y  = _v(y, "orders_promo")
        on_sf = sku in storefront_skus if storefront_skus else None
        stock_fbs = fbs_stocks_map.get(sku, {}).get("available", 0)
        stock_fby = fby_stocks_map.get(sku, {}).get("available", 0)
        stock_total = stock_fbs + stock_fby

        # ── Доля показов С БУСТОМ: вчера / ср.7д ──
        # Берём ГОТОВЫЙ процент из API (shows_boost_pct). Раньше тут считалось
        # shows_promo/shows — т.е. доля показов С АКЦИЕЙ, подписанная как «буст».
        bshow_y = _vnone(y, "shows_boost_pct")
        bshow_avg = rolling_avg(sku, "shows_boost_pct")
        bshow_avg = round(bshow_avg, 1) if bshow_avg else None

        corr_stock.append(stock_total)
        corr_shows.append(shows_y)
        corr_clicks.append(clicks_y)
        corr_carts.append(carts_y)
        corr_orders.append(orders_y)

        if on_sf is False and shows_y > 0:
            status = "🔴 Нет на витрине"
        elif on_sf is True and shows_y == 0:
            status = "🔴 0 показов"
        elif shows_avg > 0 and shows_y < shows_avg * 0.5:
            status = "🟡 Падение"
        elif on_sf is None:
            status = "—"
        elif shows_db > 0 and shows_y >= shows_db * 1.2:
            status = "🟢 Рост"
        else:
            status = "🟢 OK"

        row_data = [
            name[:50], sku, int(stock_fbs), int(stock_fby),
            int(shows_y),
            int(shows_db), int(shows_w), round(shows_avg, 1),
            _delta(shows_y, shows_db), _delta(shows_y, shows_w),
            f"{bshow_y}%" if bshow_y is not None else "—",
            f"{bshow_avg}%" if bshow_avg is not None else "—",
            int(clicks_y), int(clicks_db),
            f"{_pct(clicks_y, shows_y)}%" if _pct(clicks_y, shows_y) else "—",
            f"{_pct(clicks_db, shows_db)}%" if _pct(clicks_db, shows_db) else "—",
            int(carts_y), int(carts_db),
            f"{_pct(clicks_y, shows_y)}%" if _pct(clicks_y, shows_y) else "—",
            f"{_pct(clicks_db, shows_db)}%" if _pct(clicks_db, shows_db) else "—",
            f"{_pct(carts_y, clicks_y)}%" if _pct(carts_y, clicks_y) else "—",
            f"{_pct(carts_db, clicks_db)}%" if _pct(carts_db, clicks_db) else "—",
            int(orders_y), int(orders_db),
            int(boost_y),
            f"{_pct(boost_y, orders_y)}%" if _pct(boost_y, orders_y) else "—",
            status,
        ]
        ws2.append(row_data)

        # Color code row
        r = ws2.max_row
        for c in range(1, len(heads) + 1):
            cell = ws2.cell(row=r, column=c)
            cell.border = _border()
            cell.alignment = _center() if c > 2 else _left()

        if "🔴" in status:
            for c in range(1, len(heads) + 1):
                ws2.cell(row=r, column=c).fill = _fill(RED_LIGHT)
        elif "🟡" in status:
            for c in range(1, len(heads) + 1):
                ws2.cell(row=r, column=c).fill = _fill(YELLOW_LT)

        # Color stock cells (FBS, FBY отдельно): 0 остаток при показах → красный флаг
        for col_idx, stock_val in ((COL_STOCK_FBS, stock_fbs), (COL_STOCK_FBY, stock_fby)):
            stock_cell = ws2.cell(row=r, column=col_idx)
            if stock_val == 0 and shows_y > 0:
                stock_cell.fill = _fill(RED_LIGHT)
                stock_cell.font = _tfont(bold=True, color="B71C1C")
            elif stock_total > 0 and shows_y == 0:
                stock_cell.fill = _fill(ORANGE_LT)
                stock_cell.font = _tfont(bold=True, color="E65100")
            elif stock_val > 0:
                stock_cell.fill = _fill(GREEN_LT)

        # Color delta cells
        for col_idx in [COL_DELTA1, COL_DELTA2]:
            cell = ws2.cell(row=r, column=col_idx)
            val = str(cell.value)
            if val.startswith("+"):
                cell.font = _tfont(bold=True, color="1B5E20")
            elif val.startswith("-"):
                cell.font = _tfont(bold=True, color="B71C1C")

        # Heatmap boost share of shows (вчера): more boost = more orange
        if bshow_y is not None:
            bc = ws2.cell(row=r, column=COL_BSHOW_Y)
            if bshow_y >= 50:
                bc.fill = _fill(ORANGE_LT); bc.font = _tfont(bold=True, color="E65100")
            elif bshow_y > 0:
                bc.fill = _fill(YELLOW_LT)

        # Heatmap for Shows, Clicks, Carts, Orders
        # Shows: вчера vs позавчера (наглядное сравнение дней), позавчера vs ср.7д
        heatmap_pairs = [
            (COL_SHOWS_Y,   shows_y,   shows_db),
            (COL_SHOWS_DB,  shows_db,  shows_avg),
            (COL_CLICKS_Y,  clicks_y,  clicks_avg),
            (COL_CLICKS_DB, clicks_db, clicks_avg),
            (COL_CARTS_Y,   carts_y,   carts_avg),
            (COL_CARTS_DB,  carts_db,  carts_avg),
            (COL_ORDERS_Y,  orders_y,  orders_avg),
            (COL_ORDERS_DB, orders_db, orders_avg),
        ]
        for col_idx, val, avg in heatmap_pairs:
            if avg > 0:
                cell = ws2.cell(row=r, column=col_idx)
                if val > avg * 1.3:
                    cell.fill = _fill(GREEN_LT)
                    cell.font = _tfont(bold=True, color="1B5E20")
                elif val == 0:
                    cell.fill = _fill(RED_LIGHT)
                    cell.font = _tfont(bold=True, color="B71C1C")
                elif val < avg * 0.5:
                    cell.fill = _fill(YELLOW_LT)
                    cell.font = _tfont(bold=True, color="F57F17")

    # ── Correlation summary row ──
    def _pearson(x, y_arr):
        a, b = np.array(x, dtype=float), np.array(y_arr, dtype=float)
        if len(a) < 3 or np.std(a) == 0 or np.std(b) == 0:
            return None
        return round(float(np.corrcoef(a, b)[0, 1]), 2)

    r_shows  = _pearson(corr_stock, corr_shows)
    r_clicks = _pearson(corr_stock, corr_clicks)
    r_carts  = _pearson(corr_stock, corr_carts)
    r_orders = _pearson(corr_stock, corr_orders)

    # Empty row + correlation header
    ws2.append([])
    cr_row = ws2.max_row + 1
    ws2.cell(row=cr_row, column=1, value="📊 Корреляция остаток ↔ метрики (по всем SKU)")
    ws2.cell(row=cr_row, column=1).font = _tfont(sz=11, bold=True)
    ws2.merge_cells(start_row=cr_row, start_column=1, end_row=cr_row, end_column=2)

    corr_labels = [
        (COL_SHOWS_Y, "Показы", r_shows),
        (COL_CLICKS_Y, "Клики", r_clicks),
        (COL_CARTS_Y, "Корзина", r_carts),
        (COL_ORDERS_Y, "Заказы", r_orders),
    ]
    for col, label, r_val in corr_labels:
        cell_lbl = ws2.cell(row=cr_row, column=col, value=label)
        cell_lbl.font = _tfont(sz=10, bold=True)
        cell_lbl.alignment = _center()
        cell_lbl.border = _border()
        cell_lbl.fill = _fill(GREY_LT)

        r_text = f"r={r_val}" if r_val is not None else "—"
        cell_val = ws2.cell(row=cr_row, column=col + 1, value=r_text)
        cell_val.alignment = _center()
        cell_val.border = _border()
        cell_val.font = _tfont(bold=True)
        if r_val is not None:
            if r_val >= 0.5:
                cell_val.fill = _fill(GREEN_LT)
                cell_val.font = _tfont(bold=True, color="1B5E20")
            elif r_val <= -0.3:
                cell_val.fill = _fill(RED_LIGHT)
                cell_val.font = _tfont(bold=True, color="B71C1C")
            else:
                cell_val.fill = _fill(YELLOW_LT)

    ws2.column_dimensions["A"].width = 42
    ws2.column_dimensions["B"].width = 16
    for i in range(3, len(heads) + 1):
        ws2.column_dimensions[get_column_letter(i)].width = 14
    ws2.freeze_panes = "E2"
    ws2.auto_filter.ref = f"A1:{get_column_letter(len(heads))}{ws2.max_row}"

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 3: С АКЦИЯМИ vs БЕЗ АКЦИЙ (за вчера)
    # ⚠️ Лист назывался «Буст vs Органика» и это была ОШИБКА ПОДПИСИ: данные берутся
    # из колонок отчёта «Показы моих товаров С АКЦИЯМИ», «Клики по товарам с акциями»,
    # «Добавления в корзину ПО АКЦИЯМ» — это АКЦИИ, а не буст. Реальную долю буста
    # Яндекс отдаёт только процентом («Доля показов с бустом, %» → shows_boost_pct),
    # абсолютных чисел по бусту в этом отчёте нет.
    # ══════════════════════════════════════════════════════════════════════════
    ws3 = wb.create_sheet("Акции vs Без акций")
    ws3.sheet_properties.tabColor = "E67E22"
    boost_heads = [
        "Товар", "SKU",
        "Заказы\nвсего", "С акцией", "% акций",
        "Показы\nвсего", "Показы\nс акцией", "Показы\nбез акции", "% показов\nс акцией",
        "Корзина\nвсего", "Корзина\nс акцией", "% корзины\nс акцией",
        "Доля показов\nс БУСТОМ %",
    ]
    ws3.append(boost_heads)
    _style_header_row(ws3, 1, len(boost_heads), "E67E22")
    ws3.row_dimensions[1].height = 36
    ws3.cell(row=1, column=13).comment = Comment(
        "Единственная колонка на этом листе, которая действительно про БУСТ.\n"
        "Яндекс отдаёт её только процентом («Доля показов с бустом, %»),\n"
        "абсолютных показов с буста в отчёте нет — поэтому остальные колонки\n"
        "листа про АКЦИИ, а не про буст.", "агент")

    # Колонки "с акцией" / "без акции (показы)" — подсветка для наглядности
    BOOST_COLS = [4, 7, 11]    # С акцией, Показы с акцией, Корзина с акцией
    ORGANIC_COLS = [8]         # Показы без акции
    for c in BOOST_COLS:
        ws3.cell(row=1, column=c).fill = _fill(ORANGE_BG)
    for c in ORGANIC_COLS:
        ws3.cell(row=1, column=c).fill = _fill(GREEN_BG)

    for offer in active_offers:
        sku  = offer.get("shopSku", "")
        name = offer.get("name", sku)
        y = y_idx.get(sku, {})
        orders   = _v(y, "orders"); boost   = _v(y, "orders_promo")
        shows    = _v(y, "shows");  shows_b = _v(y, "shows_promo")
        carts    = _v(y, "carts");  carts_b = _v(y, "carts_promo")

        row_data = [
            name[:50], sku,
            int(orders), int(boost),
            f"{_pct(boost, orders)}%" if _pct(boost, orders) else "—",
            int(shows), int(shows_b), int(shows - shows_b),
            f"{_pct(shows_b, shows)}%" if _pct(shows_b, shows) else "—",
            int(carts), int(carts_b),
            f"{_pct(carts_b, carts)}%" if _pct(carts_b, carts) else "—",
            # реальная доля буста — приходит из API готовым процентом
            (f"{_v(y, 'shows_boost_pct'):.1f}%" if _vnone(y, "shows_boost_pct") is not None else "—"),
        ]
        ws3.append(row_data)
        r = ws3.max_row
        for c in range(1, len(boost_heads) + 1):
            cell = ws3.cell(row=r, column=c)
            cell.border = _border()
            cell.alignment = _center() if c > 2 else _left()

        for c in BOOST_COLS:
            ws3.cell(row=r, column=c).fill = _fill(ORANGE_LT)
        for c in ORGANIC_COLS:
            ws3.cell(row=r, column=c).fill = _fill(GREEN_LT)

    ws3.column_dimensions["A"].width = 42
    ws3.column_dimensions["B"].width = 16
    for i in range(3, len(boost_heads) + 1):
        ws3.column_dimensions[get_column_letter(i)].width = 14
    ws3.auto_filter.ref = f"A1:{get_column_letter(len(boost_heads))}{ws3.max_row}"

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 4: FBS / FBY — Продажи + Остатки + Прогноз поставки
    # ══════════════════════════════════════════════════════════════════════════
    ws4 = wb.create_sheet("FBS FBY Остатки")
    ws4.sheet_properties.tabColor = "8E44AD"

    # Build name lookup from all_offers
    name_map = {o["shopSku"]: o.get("name", o["shopSku"]) for o in all_offers}

    fbs_orders = orders_by_type.get("fbs", {})
    fby_orders = orders_by_type.get("fby", {})
    fbs_stocks = stocks.get("fbs", {})
    fby_stocks = stocks.get("fby", {})

    # Collect all SKUs that have any activity
    all_skus = sorted(
        set(fbs_stocks.keys()) | set(fby_stocks.keys())
        | {sku for day in fbs_orders.values() for sku in day}
        | {sku for day in fby_orders.values() for sku in day}
    )

    # Helper: sum orders for a sku on a specific date
    def day_orders(orders_dict, sku, dt: date):
        return orders_dict.get(dt.isoformat(), {}).get(sku, 0)

    # Title
    ws4.merge_cells("A1:P1")
    ws4["A1"] = f"FBS / FBY — Продажи и остатки · {run_date.strftime('%d.%m.%Y')}"
    ws4["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    ws4["A1"].fill = _fill("8E44AD")
    ws4["A1"].alignment = Alignment(horizontal="center")
    ws4.row_dimensions[1].height = 36

    row4 = 3
    heads4 = [
        "Товар", "SKU",
        f"FBS зак.\n{yesterday.strftime('%d.%m')}", f"FBS зак.\n{day_before.strftime('%d.%m')}",
        "FBS\nостаток", "FBS\nдоступно",
        f"FBY зак.\n{yesterday.strftime('%d.%m')}", f"FBY зак.\n{day_before.strftime('%d.%m')}",
        "FBY\nостаток", "FBY\nдоступно", "FBY\nкарантин",
        "FBY ср.\nпродаж/д", "FBY хватит\nна дней",
        "FBY нужно\nзаказать", "Срочность",
    ]
    for c, h in enumerate(heads4, 1):
        ws4.cell(row=row4, column=c, value=h)
    _style_header_row(ws4, row4, len(heads4), "8E44AD")
    ws4.row_dimensions[row4].height = 36

    # Index restock data by SKU
    restock_idx = {r["sku"]: r for r in fby_restock}

    for sku in all_skus:
        # Skip SKUs with zero FBY stock and zero FBY orders yesterday & day before
        fby_fit = fby_stocks.get(sku, {}).get("fit", 0)
        fby_ord_y = day_orders(fby_orders, sku, yesterday)
        fby_ord_db = day_orders(fby_orders, sku, day_before)
        if fby_fit == 0 and fby_ord_y == 0 and fby_ord_db == 0:
            continue

        name = name_map.get(sku, sku)
        fbs_st = fbs_stocks.get(sku, {})
        fby_st = fby_stocks.get(sku, {})
        rst = restock_idx.get(sku, {})

        row4 += 1
        vals = [
            name[:50], sku,
            day_orders(fbs_orders, sku, yesterday), day_orders(fbs_orders, sku, day_before),
            fbs_st.get("fit", 0), fbs_st.get("available", 0),
            day_orders(fby_orders, sku, yesterday), day_orders(fby_orders, sku, day_before),
            fby_st.get("fit", 0), fby_st.get("available", 0), fby_st.get("quarantine", 0),
            rst.get("avg_daily", 0), rst.get("days_left", "—"),
            rst.get("restock_qty", 0), rst.get("urgency", "—"),
        ]
        for c, v in enumerate(vals, 1):
            cell = ws4.cell(row=row4, column=c, value=v)
            cell.border = _border()
            cell.alignment = _center() if c > 2 else _left()

        # Color code urgency
        urgency = str(rst.get("urgency", ""))
        if "🔴" in urgency:
            for c in range(1, len(heads4) + 1):
                ws4.cell(row=row4, column=c).fill = _fill(RED_LIGHT)
        elif "🟡" in urgency:
            for c in range(1, len(heads4) + 1):
                ws4.cell(row=row4, column=c).fill = _fill(ORANGE_LT)
        elif "🔵" in urgency:
            for c in range(1, len(heads4) + 1):
                ws4.cell(row=row4, column=c).fill = _fill(BLUE_LT)

    ws4.column_dimensions["A"].width = 42
    ws4.column_dimensions["B"].width = 16
    for i in range(3, len(heads4) + 1):
        ws4.column_dimensions[get_column_letter(i)].width = 14
    ws4.freeze_panes = "C4"
    ws4.auto_filter.ref = f"A3:{get_column_letter(len(heads4))}{row4}"

    # ══════════════════════════════════════════════════════════════════════════
    # ЛИСТ: FBY остатки по дням (накопленная история — API её не даёт)
    # ══════════════════════════════════════════════════════════════════════════
    fby_hist = kwargs.get("fby_stock_history", {}) or {}
    if fby_hist:
        wsf = wb.create_sheet("FBY остатки по дням")
        wsf.sheet_properties.tabColor = "6C3483"
        hist_dates = sorted(fby_hist.keys())          # хронологически
        disp_dates = list(reversed(hist_dates))       # свежая слева
        # SKU, у которых был хоть какой-то остаток за период
        hist_skus = sorted({sku for day in fby_hist.values() for sku in day})

        note = ("Динамика остатков FBY (склад Яндекса) по дням — накапливается "
                "каждым прогоном отчёта (в API истории нет, только текущий снимок). "
                "Показан физический остаток (fit). «Δ вчера» = изменение за последние "
                "сутки (зелёное — приход/поставка, красное — расход). Пустая ячейка = "
                "нет данных за тот день (отчёт не собирался).")
        ncol_f = 3 + len(disp_dates)
        wsf.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol_f)
        wsf.cell(row=1, column=1, value=note).font = _tfont(sz=10, color="555555")
        wsf.cell(row=1, column=1).alignment = _left()
        wsf.row_dimensions[1].height = 46

        headsf = ["Товар", "SKU", "Δ вчера"] + [
            date.fromisoformat(d).strftime("%d.%m") for d in disp_dates]
        for c, h in enumerate(headsf, 1):
            wsf.cell(row=2, column=c, value=h)
        _style_header_row(wsf, 2, len(headsf), "6C3483")
        wsf.row_dimensions[2].height = 30

        rf = 2
        for sku in hist_skus:
            rf += 1
            # значения по дням (свежая слева)
            vals = []
            for d in disp_dates:
                cell = fby_hist.get(d, {}).get(sku)
                vals.append(cell["fit"] if cell else None)
            # Δ за последние сутки: два самых свежих дня, где есть значение
            delta = None
            recent_nonnull = [v for v in vals if v is not None]
            if len(vals) >= 2 and vals[0] is not None and vals[1] is not None:
                delta = vals[0] - vals[1]
            wsf.cell(row=rf, column=1, value=name_map.get(sku, sku)[:50])
            wsf.cell(row=rf, column=2, value=sku)
            dc = wsf.cell(row=rf, column=3,
                          value=(f"{delta:+d}" if delta is not None else "—"))
            if delta is not None and delta > 0:
                dc.font = _tfont(bold=True, color="1B5E20")
            elif delta is not None and delta < 0:
                dc.font = _tfont(bold=True, color="B71C1C")
            for i, v in enumerate(vals):
                cell = wsf.cell(row=rf, column=4 + i, value=(v if v is not None else ""))
                # 0 остатка при наличии данных — красный флаг
                if v == 0:
                    cell.fill = _fill(RED_LIGHT)
                    cell.font = _tfont(bold=True, color="B71C1C")
                elif v is not None and v > 0:
                    cell.fill = _fill(GREEN_LT)
            for c in range(1, len(headsf) + 1):
                cell = wsf.cell(row=rf, column=c)
                cell.border = _border()
                cell.alignment = _center() if c > 2 else _left()

        wsf.column_dimensions["A"].width = 40
        wsf.column_dimensions["B"].width = 14
        wsf.column_dimensions["C"].width = 9
        for i in range(4, len(headsf) + 1):
            wsf.column_dimensions[get_column_letter(i)].width = 7
        wsf.freeze_panes = "D3"
        wsf.auto_filter.ref = f"A2:{get_column_letter(len(headsf))}{rf}"

    # ══════════════════════════════════════════════════════════════════════════
    # ЛИСТ: Ставки буста по дням (вся история из boost_bids_daily.csv, без обрезки)
    # ══════════════════════════════════════════════════════════════════════════
    bids_hist = kwargs.get("boost_bids_history", {}) or {}
    promo_hist = kwargs.get("promo_history", {}) or {}
    promo_meta = kwargs.get("promo_meta", {}) or {}
    # Достаточно ЛЮБОЙ из двух историй: пока буст ведётся кампаниями кабинета,
    # bids/info пуст — раньше при пустых ставках лист исчезал бы вместе с акциями.
    if bids_hist or promo_hist:
        wsr = wb.create_sheet("Ставки буста по дням")
        wsr.sheet_properties.tabColor = "AF601A"
        # даты: объединяем историю ставок и историю акций (могли начаться в разные дни)
        rdates = sorted(set(bids_hist.keys()) | set(promo_hist.keys()))
        rdisp = list(reversed(rdates))                # свежая слева
        rskus = sorted({sku for day in bids_hist.values() for sku in day}
                       | {sku for day in promo_hist.values() for sku in day})

        # Короткие коды акций (А1, А2…) — полное название в колонку не влезет.
        # Порядок стабильный: по дате старта, затем по названию.
        # ЗНАЧЕНИЕ в promo_hist — СПИСОК id (артикул часто в НЕСКОЛЬКИХ акциях
        # одновременно, проверено на живых данных: 14 из 26 активных SKU сразу
        # в двух). Раньше тут стоял единичный id — вторая акция терялась.
        promo_ids = sorted({pid for day in promo_hist.values()
                            for pids in day.values() for pid in (pids or [])},
                           key=lambda p: ((promo_meta.get(p) or {}).get("from", ""),
                                          (promo_meta.get(p) or {}).get("name", ""), p))
        promo_code = {pid: f"А{i}" for i, pid in enumerate(promo_ids, 1)}
        PROMO_TINTS = ("C6EFCE", "BDD7EE", "FFE0B2", "E8DAEF", "FFF9C4", "D5F5E3")
        promo_tint = {pid: PROMO_TINTS[i % len(PROMO_TINTS)]
                      for i, pid in enumerate(promo_ids)}
        legend = "; ".join(
            f"{promo_code[p]} = {(promo_meta.get(p) or {}).get('name', p)[:38]}"
            f" ({(promo_meta.get(p) or {}).get('from', '?')}→{(promo_meta.get(p) or {}).get('to', '?')})"
            for p in promo_ids)

        rnote = ("История API-ставок буста по дням (из boost_bids_daily.csv, "
                 "накапливается каждым прогоном, без обрезки). «Δ вчера» — изменение "
                 "ставки за сутки. Жёлтая ячейка — день, когда ставка изменилась "
                 "относительно предыдущего. Пусто = SKU не было в API-ставках в тот "
                 "день (ведётся групповой кампанией кабинета или буст не задан). "
                 "Под каждым артикулом строка «Акция» — в каких акциях он был в этот "
                 "день (из promo_daily.json), артикул может быть в НЕСКОЛЬКИХ акциях "
                 "одновременно — тогда коды через запятую (напр. «А1,А2»). ВАЖНО: "
                 "истории участия в акциях у Яндекса в API НЕТ, она накапливается только "
                 "с первого прогона с этой правкой — за более ранние дни строка пустая, "
                 "это не ошибка. "
                 + (f"АКЦИИ: {legend}" if legend else ""))
        ncol_r = 4 + len(rdisp)
        wsr.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol_r)
        wsr.cell(row=1, column=1, value=rnote).font = _tfont(sz=10, color="555555")
        wsr.cell(row=1, column=1).alignment = _left()
        wsr.row_dimensions[1].height = 74

        headsr = ["Товар", "SKU", "Строка", "Δ вчера"] + [
            date.fromisoformat(d).strftime("%d.%m") for d in rdisp]
        DAY0 = 5   # первая колонка дня
        for c, h in enumerate(headsr, 1):
            wsr.cell(row=2, column=c, value=h)
        _style_header_row(wsr, 2, len(headsr), "AF601A")
        wsr.row_dimensions[2].height = 30

        rr = 2
        for sku in rskus:
            # ── строка 1: ставка буста ──
            rr += 1
            vals = [bids_hist.get(d, {}).get(sku) for d in rdisp]
            delta = None
            if len(vals) >= 2 and isinstance(vals[0], (int, float)) and isinstance(vals[1], (int, float)):
                delta = round(vals[0] - vals[1], 2)
            wsr.cell(row=rr, column=1, value=name_map.get(sku, sku)[:50])
            wsr.cell(row=rr, column=2, value=sku)
            sc = wsr.cell(row=rr, column=3, value="Ставка %")
            sc.font = _tfont(bold=True, color="B9770E"); sc.fill = _fill(ORANGE_LT)
            dc = wsr.cell(row=rr, column=4,
                          value=(f"{delta:+g}" if delta is not None else "—"))
            if delta:
                dc.font = _tfont(bold=True, color="1B5E20" if delta > 0 else "B71C1C")
            for i, v in enumerate(vals):
                cell = wsr.cell(row=rr, column=DAY0 + i,
                                value=(v if isinstance(v, (int, float)) else ""))
                if isinstance(v, (int, float)):
                    cell.number_format = '0.0"%"'
                    # день изменения ставки (относительно следующего справа = предыдущего дня)
                    prev_v = vals[i + 1] if i + 1 < len(vals) else None
                    if isinstance(prev_v, (int, float)) and abs(v - prev_v) > 1e-9:
                        cell.fill = _fill(YELLOW_LT)
                        cell.font = _tfont(bold=True)
            for c in range(1, len(headsr) + 1):
                cell = wsr.cell(row=rr, column=c)
                cell.border = _border()
                cell.alignment = _center() if c > 2 else _left()

            # ── строка 2: участие в акции(ях) по дням ──
            rr += 1
            # каждый элемент — СПИСОК promo_id за этот день (может быть несколько
            # одновременно), сортируем для стабильного отображения и сравнения
            pvals = [sorted(promo_hist.get(d, {}).get(sku) or []) for d in rdisp]
            wsr.cell(row=rr, column=2, value=sku)
            pc = wsr.cell(row=rr, column=3, value="Акция")
            pc.font = _tfont(bold=True, color="1B5E20"); pc.fill = _fill(GREEN_LT)
            # Δ по акции — день, когда набор акций сменился/начался/закончился
            # (сравниваем множества, не порядок — он уже стабилен после sorted())
            pdelta = "—"
            if len(pvals) >= 2 and pvals[0] != pvals[1]:
                pdelta = "смена"
            pdc = wsr.cell(row=rr, column=4, value=pdelta)
            if pdelta != "—":
                pdc.font = _tfont(bold=True, color="B9770E")
            for i, pids in enumerate(pvals):
                codes = [promo_code.get(p, "?") for p in pids]
                cell = wsr.cell(row=rr, column=DAY0 + i, value=",".join(codes))
                if pids:
                    # фон — по первой (по сортировке) акции; при 2+ акциях сразу
                    # видно из текста ячейки ("А1,А2"), фон только ориентир
                    cell.fill = _fill(promo_tint.get(pids[0], GREEN_LT))
                    cell.font = _tfont(bold=True)
                    # день смены набора акций — рамкой не выделить, помечаем курсивом
                    prev_pids = pvals[i + 1] if i + 1 < len(pvals) else []
                    if prev_pids != pids:
                        cell.font = Font(bold=True, italic=True, size=11)
            for c in range(1, len(headsr) + 1):
                cell = wsr.cell(row=rr, column=c)
                cell.border = _border()
                cell.alignment = _center() if c > 2 else _left()

        wsr.column_dimensions["A"].width = 40
        wsr.column_dimensions["B"].width = 14
        wsr.column_dimensions["C"].width = 10
        wsr.column_dimensions["D"].width = 9
        for i in range(DAY0, len(headsr) + 1):
            wsr.column_dimensions[get_column_letter(i)].width = 9  # хватает под "А1,А2"
        wsr.freeze_panes = "E3"
        wsr.auto_filter.ref = f"A2:{get_column_letter(len(headsr))}{rr}"

    # ══════════════════════════════════════════════════════════════════════════
    # SHEETS 5-9: 14-дневная динамика по SKU
    # Показы / Клики / CTR / Покупки / CR клик→покупка
    # ══════════════════════════════════════════════════════════════════════════

    # Prepare 14-day date list (sorted ascending)
    all_dates = sorted(df_full["date"].unique()) if "date" in df_full.columns else []
    all_dates = [d for d in all_dates if d <= yesterday][-14:]  # last 14 days
    last_7 = all_dates[-7:] if len(all_dates) >= 7 else all_dates

    # Pivot: rows=shopSku, columns=dates, values=metric
    def pivot_metric(metric):
        if df_full.empty or metric not in df_full.columns or "shopSku" not in df_full.columns:
            return pd.DataFrame()
        sub = df_full[df_full["date"].isin(all_dates)][["shopSku", "date", metric]].copy()
        sub[metric] = sub[metric].fillna(0)
        return sub.pivot_table(index="shopSku", columns="date", values=metric, aggfunc="sum", fill_value=0)

    pv_shows  = pivot_metric("shows")
    pv_clicks = pivot_metric("clicks")
    pv_orders = pivot_metric("orders")
    pv_carts  = pivot_metric("carts")
    pv_shows_promo = pivot_metric("shows_promo")
    # «без акции» = все показы − показы с акцией (по дням), выравниваем индексы
    if not pv_shows.empty:
        pv_shows_org = pv_shows.subtract(
            pv_shows_promo.reindex(index=pv_shows.index, columns=pv_shows.columns, fill_value=0),
            fill_value=0).clip(lower=0)
    else:
        pv_shows_org = pd.DataFrame()

    # Name lookup
    name_map_full = {o["shopSku"]: o.get("name", o["shopSku"]) for o in all_offers}

    # Filter to active SKUs only (those with any shows in 14 days)
    active_skus = set()
    if not pv_shows.empty:
        active_skus = set(pv_shows[pv_shows.sum(axis=1) > 0].index)

    def _trend_slope(values):
        """Linear regression slope over values. Positive = growing."""
        if len(values) < 2:
            return 0
        x = list(range(len(values)))
        try:
            slope = np.polyfit(x, values, 1)[0]
            return round(slope, 2)
        except Exception:
            return 0

    def _trend_label(slope, avg):
        """Human-readable trend label with direction."""
        if avg == 0:
            return "—"
        # slope per day as % of average
        pct = abs(slope / avg * 100) if avg else 0
        if slope > 0 and pct > 5:
            return f"↗ +{round(pct)}%/д"
        elif slope < 0 and pct > 5:
            return f"↘ -{round(pct)}%/д"
        else:
            return "→ стабильно"

    def _trend_label_resid(slope, scale):
        """Как _trend_label, но для тренда по ОСТАТКАМ (факт − норма для объёма):
        slope уже в п.п./день, scale — средний уровень метрики (для %), чтобы
        выразить сдвиг в тех же % от масштаба. «К норме» означает: это реальный
        сдвиг качества, объём уже вычтен."""
        if not scale:
            return "→ по норме"
        pct = abs(slope / scale * 100)
        if slope > 0 and pct > 5:
            return f"↗ +{round(pct)}%/д (к норме)"
        elif slope < 0 and pct > 5:
            return f"↘ -{round(pct)}%/д (к норме)"
        else:
            return "→ по норме"

    TREND_UP_COLOR   = "1B5E20"  # dark green
    TREND_DOWN_COLOR = "B71C1C"  # dark red
    TREND_FLAT_COLOR = "555555"  # grey

    def add_dynamics_sheet(wb, sheet_name, tab_color, pv_data, metric_label,
                          is_ratio=False, numerator_pv=None, denominator_pv=None,
                          add_period_sum=False, stats_first=False):
        """
        Add a 14-day dynamics sheet.
        If is_ratio=True, compute values as numerator/denominator * 100 per cell.
        If stats_first=True: колонки после SKU — Медиана/Среднее, затем даты по
        убыванию (свежие слева); Сумма-за-период в этом режиме не поддерживается.
        """
        ws = wb.create_sheet(sheet_name)
        ws.sheet_properties.tabColor = tab_color

        want_sum = add_period_sum and not is_ratio and not stats_first

        # Header row
        if stats_first:
            headers = ["Товар", "SKU", "Медиана\n14д", "Среднее\n14д"]
            for d in reversed(all_dates):
                headers.append(d.strftime("%d.%m"))
            headers.append("Тренд 7д")
            date_start_col = 5
        else:
            headers = ["Товар", "SKU"]
            for d in all_dates:
                headers.append(d.strftime("%d.%m"))
            if want_sum:
                headers.append("Сумма\nза 14д")
            if is_ratio:
                headers += ["Медиана\n14д", "Среднее\nвзвеш. 14д", "Тренд 7д\n(к норме объёма)"]
            else:
                headers += ["Медиана\n14д", "Среднее\n14д", "Тренд 7д"]
            date_start_col = 3
        col_sum = 3 + len(all_dates) if want_sum else None

        ws.append(headers)
        _style_header_row(ws, 1, len(headers), tab_color)
        ws.row_dimensions[1].height = 36

        # НОРМА ДЛЯ ОБЪЁМА (только для долевых метрик, CTR/CR): доля мех. падает в
        # дни с необычно большим объёмом знаменателя (показов/кликов) — широкий
        # трафик разбавляет целевую аудиторию. Без поправки такой день ложно читается
        # как «просадка», хотя он в норме для СВОЕГО объёма (см. «Конверсия карточек»
        # — та же идея: rate~log(объём)). Регрессию строим КРОСС-СЕКЦИОННО ПО ДАТЕ —
        # по всем активным SKU в этот конкретный день, а не по своим 14 дням SKU:
        # у одного SKU обычно всего 1 день-выброс по объёму, и регрессия по своей же
        # истории тогда либо идеально «объясняет» этот день сам собой (0 диагностики),
        # либо экстраполирует на порядки за пределы диапазона своих обычных объёмов
        # (мусорный прогноз). Кросс-секционно по дню — у каталога всегда есть разброс
        # объёмов между SKU, так что норма для высокого объёма опирается на РЕАЛЬНЫЕ
        # чужие карточки с похожим объёмом в тот же день, а не на экстраполяцию.
        day_norm_curve = {}
        if is_ratio and numerator_pv is not None and denominator_pv is not None:
            for d in all_dates:
                dens, vals = [], []
                for s in active_skus:
                    den = denominator_pv.at[s, d] if s in denominator_pv.index and d in denominator_pv.columns else 0
                    num = numerator_pv.at[s, d] if s in numerator_pv.index and d in numerator_pv.columns else 0
                    if den > 0:
                        dens.append(float(den)); vals.append(float(num) / float(den) * 100)
                dens_arr = np.array(dens); vals_arr = np.array(vals)
                if len(dens_arr) >= 8 and float(np.ptp(np.log(dens_arr + 1))) > 0.5:
                    try:
                        b1, b0 = np.polyfit(np.log(dens_arr + 1), vals_arr, 1, w=dens_arr)
                        day_norm_curve[d] = (b1, b0)
                    except Exception:
                        pass

        for sku in sorted(active_skus):
            name = name_map_full.get(sku, sku)

            num_values = []
            den_values = []
            if is_ratio and numerator_pv is not None and denominator_pv is not None:
                # Compute ratio per day (хронологический порядок, как all_dates)
                day_values = []
                for d in all_dates:
                    num = numerator_pv.at[sku, d] if sku in numerator_pv.index and d in numerator_pv.columns else 0
                    den = denominator_pv.at[sku, d] if sku in denominator_pv.index and d in denominator_pv.columns else 0
                    num_values.append(float(num)); den_values.append(float(den))
                    val = round(num / den * 100, 2) if den > 0 else 0
                    day_values.append(val)
            else:
                day_values = []
                for d in all_dates:
                    val = pv_data.at[sku, d] if sku in pv_data.index and d in pv_data.columns else 0
                    day_values.append(int(val))

            # Норма для объёма — из кросс-секционной кривой этой даты (посчитана
            # один раз для всех SKU выше, до цикла).
            norm_vals = None
            if is_ratio and day_norm_curve:
                norm_vals = []
                for i, d in enumerate(all_dates):
                    curve = day_norm_curve.get(d)
                    if curve and den_values[i] > 0:
                        b1, b0 = curve
                        norm_vals.append(round(b0 + b1 * np.log(den_values[i] + 1), 2))
                    else:
                        norm_vals.append(None)
                if all(v is None for v in norm_vals):
                    norm_vals = None

            # Stats. Для счётных метрик (показы/клики/покупки) медиану и среднее
            # показываем целыми; для долей (CTR/CR) — среднее ВЗВЕШЕНО по знаменателю
            # (Σclicks/Σshows), а не средним из дневных %, чтобы один день с большим
            # объёмом не перекашивал цифру непропорционально своему реальному вкладу.
            if day_values:
                if is_ratio:
                    median_val = round(statistics.median(day_values), 2)
                    den_sum = sum(den_values)
                    avg_val = (round(sum(num_values) / den_sum * 100, 2) if den_sum > 0
                               else round(statistics.mean(day_values), 2))
                else:
                    median_val = round(statistics.median(day_values))
                    avg_val = round(statistics.mean(day_values))
            else:
                median_val = 0
                avg_val = 0

            # 7-дневный тренд: если норма по объёму посчитана — тренд по ОСТАТКАМ
            # (факт − норма для объёма того дня), чтобы дни-выбросы по объёму не
            # создавали ложный тренд «падает/растёт» на ровном месте.
            if is_ratio and norm_vals is not None:
                resid_vals = [(day_values[i] - norm_vals[i]) if norm_vals[i] is not None else 0.0
                              for i in range(len(day_values))]
                trend_vals = resid_vals[-7:] if len(resid_vals) >= 7 else resid_vals
                slope = _trend_slope(trend_vals)
                trend_text = _trend_label_resid(slope, avg_val)
            else:
                trend_vals = day_values[-7:] if len(day_values) >= 7 else day_values
                slope = _trend_slope(trend_vals)
                trend_avg = statistics.mean(trend_vals) if trend_vals else 0
                trend_text = _trend_label(slope, trend_avg)

            # Значения для вывода в колонки дат — в том же порядке, что и headers
            display_values = list(reversed(day_values)) if stats_first else day_values
            display_norm = ((list(reversed(norm_vals)) if stats_first else norm_vals)
                            if norm_vals is not None else None)

            if stats_first:
                row_data = [name[:50], sku, median_val, avg_val] + display_values + [trend_text]
            else:
                trailing = []
                if want_sum:
                    trailing.append(sum(day_values))
                trailing += [median_val, avg_val, trend_text]
                row_data = [name[:50], sku] + display_values + trailing
            ws.append(row_data)

            r = ws.max_row
            for c in range(1, len(headers) + 1):
                cell = ws.cell(row=r, column=c)
                cell.border = _border()
                cell.alignment = _center() if c > 2 else _left()

                # Color-code daily values (heatmap: higher = греенее). Для is_ratio
                # база сравнения — НОРМА ДЛЯ ОБЪЁМА ЭТОГО ДНЯ (если посчитана), а не
                # плоское среднее — иначе день с большим объёмом ложно красится как
                # просадка (см. _trend_label_resid выше).
                if date_start_col <= c <= date_start_col - 1 + len(all_dates) and avg_val > 0:
                    idx = c - date_start_col
                    val = display_values[idx]
                    if is_ratio:
                        cell.number_format = '0.0"%"'
                    baseline = (display_norm[idx] if (display_norm is not None
                                and display_norm[idx] is not None) else avg_val)
                    if baseline > 0 and val > baseline * 1.3:
                        cell.fill = _fill(GREEN_LT)
                    elif baseline > 0 and val < baseline * 0.5 and val > 0:
                        cell.fill = _fill(YELLOW_LT)
                    elif baseline > 0 and val == 0:
                        cell.fill = _fill(RED_LIGHT)

            # Period-sum cell styling (итог по артикулу за 14д)
            if col_sum is not None:
                sc = ws.cell(row=r, column=col_sum)
                sc.font = _tfont(bold=True)
                sc.fill = _fill(BLUE_LT)

            # Trend cell styling
            trend_cell = ws.cell(row=r, column=len(headers))
            if "↗" in trend_text:
                trend_cell.font = _tfont(bold=True, color=TREND_UP_COLOR)
                trend_cell.fill = _fill(GREEN_LT)
            elif "↘" in trend_text:
                trend_cell.font = _tfont(bold=True, color=TREND_DOWN_COLOR)
                trend_cell.fill = _fill(RED_LIGHT)
            else:
                trend_cell.font = _tfont(color=TREND_FLAT_COLOR)

        # Column widths
        ws.column_dimensions["A"].width = 42
        ws.column_dimensions["B"].width = 16
        for i in range(3, len(headers) + 1):
            ws.column_dimensions[get_column_letter(i)].width = 10
        if stats_first:
            # Медиана/Среднее (C,D) и Тренд (последняя) — пошире
            ws.column_dimensions["C"].width = 13
            ws.column_dimensions["D"].width = 13
            ws.column_dimensions[get_column_letter(len(headers))].width = 13
            ws.freeze_panes = "E2"
        else:
            # Wider for stats columns
            for i in range(len(all_dates) + 3, len(headers) + 1):
                ws.column_dimensions[get_column_letter(i)].width = 13
            ws.freeze_panes = "C2"
        ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{ws.max_row}"
        return ws

    def append_org_boost_breakdown(ws, start_row):
        """Под основной таблицей показов добавляет поартикульную разбивку:
        по каждому SKU две строки — Органика и Буст — по каждому из 14 дней."""
        ndates = len(all_dates)
        ncol = 3 + ndates + 1  # Товар, SKU, Канал, <дни>, Сумма 14д
        r = start_row
        # Заголовок секции
        ws.cell(row=r, column=1,
                value="РАЗБИВКА ПОКАЗОВ: БЕЗ АКЦИИ vs С АКЦИЕЙ по дням (по каждому артикулу)")
        ws.cell(row=r, column=1).font = _tfont(bold=True, sz=12, color="2E86C1")
        r += 1
        hdr = ["Товар", "SKU", "Канал"] + [d.strftime("%d.%m") for d in all_dates] + ["Сумма\n14д"]
        for c, h in enumerate(hdr, 1):
            ws.cell(row=r, column=c, value=h)
        _style_header_row(ws, r, ncol, "2E86C1")
        ws.row_dimensions[r].height = 30
        hdr_row = r
        for sku in sorted(active_skus):
            name = name_map_full.get(sku, sku)
            for channel, pv, tint, fcolor in (
                    ("Без акции", pv_shows_org, GREEN_LT, "1B5E20"),
                    ("С акцией", pv_shows_promo, ORANGE_LT, "B9770E")):
                r += 1
                days = []
                for d in all_dates:
                    v = pv.at[sku, d] if (not pv.empty and sku in pv.index and d in pv.columns) else 0
                    days.append(int(round(float(v))))
                row = [name[:50], sku, channel] + days + [sum(days)]
                for c, v in enumerate(row, 1):
                    cell = ws.cell(row=r, column=c, value=v)
                    cell.border = _border()
                    cell.alignment = _center() if c >= 3 else _left()
                ws.cell(row=r, column=3).fill = _fill(tint)
                ws.cell(row=r, column=3).font = _tfont(bold=True, color=fcolor)
                sc = ws.cell(row=r, column=ncol)
                sc.font = _tfont(bold=True); sc.fill = _fill(BLUE_LT)
        ws.freeze_panes = ws.freeze_panes  # сохраняем существующую заморозку
        return hdr_row

    # Sheet 5: Показы (+ поартикульная разбивка органика/буст ниже)
    _ws_shows = add_dynamics_sheet(wb, "Показы 14д", "2E86C1", pv_shows, "Показы",
                                   stats_first=True)
    if _ws_shows is not None and all_dates and active_skus:
        append_org_boost_breakdown(_ws_shows, _ws_shows.max_row + 3)

    # ══════════════════════════════════════════════════════════════════════════
    # ЛИСТ: Конверсия карточек (клик→корзина с нормой по объёму + красные флаги)
    # ══════════════════════════════════════════════════════════════════════════
    def _binom_cdf(k, n, p):
        """P(X<=k) для Binom(n,p) — без scipy. Норм-аппрокс с поправкой Йетса
        при n большом, точная сумма при малом (чтобы не зависеть от прод-питона)."""
        if n <= 0 or p <= 0:
            return 1.0
        p = min(max(p, 1e-9), 1 - 1e-9)
        if n <= 1000:
            from math import comb
            return float(sum(comb(n, i) * p**i * (1 - p)**(n - i) for i in range(int(k) + 1)))
        mu = n * p; sd = (n * p * (1 - p)) ** 0.5
        from math import erf, sqrt
        z = (k + 0.5 - mu) / (sd if sd > 0 else 1e-9)
        return 0.5 * (1 + erf(z / sqrt(2)))

    if not pv_clicks.empty and not pv_carts.empty and active_skus:
        rows_cv = []
        for sku in active_skus:
            shows = float(pv_shows.loc[sku].sum()) if sku in pv_shows.index else 0.0
            clicks = float(pv_clicks.loc[sku].sum()) if sku in pv_clicks.index else 0.0
            carts = float(pv_carts.loc[sku].sum()) if sku in pv_carts.index else 0.0
            if clicks >= 1:
                rows_cv.append([sku, shows, clicks, carts])
        if len(rows_cv) >= 5:
            arr = np.array([[r[1], r[2], r[3]] for r in rows_cv], dtype=float)
            # взвешенная регрессия cart_rate ~ log(shows+1), вес = клики
            X = np.log(arr[:, 0] + 1); yv = arr[:, 2] / arr[:, 1]; wv = arr[:, 1]
            try:
                b1, b0 = np.polyfit(X, yv, 1, w=wv)
            except Exception:
                b1, b0 = 0.0, float(np.average(yv, weights=wv))
            base_rate = float(np.average(yv, weights=wv))

            data_cv = []
            for sku, shows, clicks, carts in rows_cv:
                exp_rate = max(min(b0 + b1 * np.log(shows + 1), 0.9), 0.01)
                lost = clicks * exp_rate - carts
                pval = _binom_cdf(carts, int(round(clicks)), exp_rate)
                data_cv.append((sku, shows, clicks, carts, carts / clicks,
                                exp_rate, lost, pval))
            data_cv.sort(key=lambda x: -x[6])   # по потерянным корзинам ↓

            # порог «много показов» — медиана показов среди значимо-красных карточек;
            # красный флаг + показы выше медианы = дорогая утечка (чинить первым)
            red_shows = [d[1] for d in data_cv if (d[7] < 0.05 and d[2] >= 30)]
            hot_thr = float(np.median(red_shows)) if red_shows else float("inf")

            wc = wb.create_sheet("Конверсия карточек")
            wc.sheet_properties.tabColor = "884EA0"
            cnote = (f"Клик→корзина по карточкам за {len(all_dates)}д. Норма падает с объёмом "
                     f"(массовый трафик менее целевой), поэтому каждая карточка сравнивается "
                     f"с НОРМОЙ ДЛЯ СВОЕГО объёма показов (взвеш. регрессия). «Потерянные корзины» "
                     f"= сколько недобрала против нормы. 🔴 = статзначимо ниже нормы (биномиальный "
                     f"тест p<0.05 при кликах≥30). Базовый клик→корзина каталога: {base_rate*100:.1f}%. "
                     f"Красные — куда трафик идёт, а карточка не конвертит (чинить контент/цену/фото). "
                     f"🔥 = красный флаг + показы выше медианы красных = дорогая утечка, чинить первой.")
            heads_cv = ["Товар", "SKU", "Показы", "Клики", "Корзины",
                        "Клик→\nкорзина %", "Норма\nдля объёма %", "Потеряно\nкорзин",
                        "p-value", "Флаг"]
            nc = len(heads_cv)
            wc.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nc)
            wc.cell(row=1, column=1, value=cnote).font = _tfont(sz=10, color="555555")
            wc.cell(row=1, column=1).alignment = _left()
            wc.row_dimensions[1].height = 60
            for c, h in enumerate(heads_cv, 1):
                wc.cell(row=2, column=c, value=h)
            _style_header_row(wc, 2, nc, "884EA0")
            wc.row_dimensions[2].height = 34

            rw = 2
            for sku, shows, clicks, carts, cr, er, lost, pval in data_cv:
                rw += 1
                sig = (pval < 0.05 and clicks >= 30)
                if sig and shows >= hot_thr:
                    flag = "🔥 дорогая утечка"
                elif sig:
                    flag = "🔴 карточка"
                elif pval < 0.15 and clicks >= 15:
                    flag = "🟡"
                else:
                    flag = ""
                vals = [name_map_full.get(sku, sku)[:45], sku, int(shows), int(clicks),
                        int(carts), round(cr * 100, 1), round(er * 100, 1),
                        round(lost, 1), round(pval, 3), flag]
                for c, v in enumerate(vals, 1):
                    cell = wc.cell(row=rw, column=c, value=v)
                    cell.border = _border()
                    cell.alignment = _left() if c in (1, 10) else _center()
                    if c in (3, 4, 5) and isinstance(v, (int, float)):
                        cell.number_format = '#,##0'
                # подсветка строки по значимости недоконверсии
                if sig:
                    for c in (6, 8):
                        wc.cell(row=rw, column=c).fill = _fill(RED_LIGHT)
                        wc.cell(row=rw, column=c).font = _tfont(bold=True, color="B71C1C")
                elif lost < -3:   # заметно ЛУЧШЕ нормы — зелёным
                    wc.cell(row=rw, column=6).fill = _fill(GREEN_LT)
                    wc.cell(row=rw, column=6).font = _tfont(bold=True, color="1B5E20")
            wc.column_dimensions["A"].width = 38
            wc.column_dimensions["B"].width = 13
            for i in range(3, nc):
                wc.column_dimensions[get_column_letter(i)].width = 11
            wc.column_dimensions[get_column_letter(nc)].width = 12
            wc.freeze_panes = "C3"
            wc.auto_filter.ref = f"A2:{get_column_letter(nc)}{rw}"

    # Лист «Текущий тест» — тест vs контроль по БУСТ-показам (рядом с «Показы 14д»)
    try:
        from .config import (BOOST_TEST_LABEL, BOOST_TEST_START,
                            BOOST_TEST_SKUS, BOOST_CONTROL_SKUS)
    except Exception:
        BOOST_TEST_LABEL = BOOST_TEST_START = ""
        BOOST_TEST_SKUS = BOOST_CONTROL_SKUS = []

    if all_dates and (BOOST_TEST_SKUS or BOOST_CONTROL_SKUS):
        wst = wb.create_sheet("Текущий тест")
        wst.sheet_properties.tabColor = "8E44AD"
        wst.sheet_view.showGridLines = False
        ndates = len(all_dates)
        ncol = 2 + ndates + 1  # Группа/Товар, SKU, <дни>, Δ% посл.день

        def _gsum(skus, pv, d):
            return sum(int(round(float(pv.at[s, d])))
                       for s in skus
                       if (not pv.empty and s in pv.index and d in pv.columns))

        def _sku_day(sku, pv, d):
            return int(round(float(pv.at[sku, d]))) if (not pv.empty and sku in pv.index and d in pv.columns) else 0

        def _delta_last(days):
            if len(days) >= 2 and days[-2] > 0:
                return f"{round((days[-1] - days[-2]) / days[-2] * 100, 1)}%"
            return "—"

        # Заголовок
        wst.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
        wst["A1"] = (f"ТЕКУЩИЙ ТЕСТ БУСТА · {BOOST_TEST_LABEL or '—'}. "
                     f"Сравнение по БУСТ-показам (не суммарным: показы = органика + буст; "
                     f"органика — шумный остаток, может симметрично падать по внешним причинам). "
                     f"Тест = где меняли ставку, Контроль = где оставили. Δ% — последний день vs предыдущий.")
        wst["A1"].font = _tfont(sz=10, color="555555"); wst["A1"].alignment = _left()
        r = 2

        # ── Блок 1: сводка по группам (буст-показы и органика по дням) ──
        wst.cell(row=r, column=1, value="СВОДКА ПО ГРУППАМ (показы/день)").font = _tfont(bold=True, sz=12, color="8E44AD")
        r += 1
        hdr = ["Группа", "Канал"] + [d.strftime("%d.%m") for d in all_dates] + ["Δ% посл."]
        for c, h in enumerate(hdr, 1):
            wst.cell(row=r, column=c, value=h)
        _style_header_row(wst, r, ncol, "8E44AD")
        wst.row_dimensions[r].height = 28
        summary = [
            ("ТЕСТ", "С акцией", BOOST_TEST_SKUS, pv_shows_promo, ORANGE_LT),
            ("КОНТРОЛЬ", "С акцией", BOOST_CONTROL_SKUS, pv_shows_promo, ORANGE_LT),
            ("ТЕСТ", "Без акции", BOOST_TEST_SKUS, pv_shows_org, GREEN_LT),
            ("КОНТРОЛЬ", "Без акции", BOOST_CONTROL_SKUS, pv_shows_org, GREEN_LT),
        ]
        for grp, channel, skus, pv, tint in summary:
            r += 1
            days = [_gsum(skus, pv, d) for d in all_dates]
            row = [grp, channel] + days + [_delta_last(days)]
            for c, v in enumerate(row, 1):
                cell = wst.cell(row=r, column=c, value=v)
                cell.border = _border()
                cell.alignment = _left() if c == 1 else _center()
                if c == 1:
                    cell.font = _tfont(bold=True)
            wst.cell(row=r, column=2).fill = _fill(tint)
            wst.cell(row=r, column=2).font = _tfont(bold=True)

        # ── Блоки 2 и 3: по артикулам, буст-показы по дням ──
        for grp_name, skus, color in (("ТЕСТ — буст-показы по артикулам", BOOST_TEST_SKUS, "8E44AD"),
                                       ("КОНТРОЛЬ — буст-показы по артикулам", BOOST_CONTROL_SKUS, "7F8C8D")):
            r += 2
            wst.cell(row=r, column=1, value=grp_name).font = _tfont(bold=True, sz=12, color=color)
            r += 1
            hdr = ["Товар", "SKU"] + [d.strftime("%d.%m") for d in all_dates] + ["Δ% посл."]
            for c, h in enumerate(hdr, 1):
                wst.cell(row=r, column=c, value=h)
            _style_header_row(wst, r, ncol, color)
            for sku in skus:
                r += 1
                name = name_map_full.get(sku, sku)
                days = [_sku_day(sku, pv_shows_promo, d) for d in all_dates]
                row = [name[:46], sku] + days + [_delta_last(days)]
                for c, v in enumerate(row, 1):
                    cell = wst.cell(row=r, column=c, value=v)
                    cell.border = _border()
                    cell.alignment = _left() if c <= 2 else _center()

        wst.column_dimensions["A"].width = 40
        wst.column_dimensions["B"].width = 16
        for i in range(3, ncol + 1):
            wst.column_dimensions[get_column_letter(i)].width = 9
        wst.column_dimensions[get_column_letter(ncol)].width = 11
        wst.freeze_panes = "C2"

    # Sheet 6: Клики
    add_dynamics_sheet(wb, "Клики 14д", "2980B9", pv_clicks, "Клики", add_period_sum=True)

    # Sheet 7: CTR (показы → клики)
    add_dynamics_sheet(wb, "CTR 14д", "1ABC9C", None, "CTR%",
                       is_ratio=True, numerator_pv=pv_clicks, denominator_pv=pv_shows)

    # Sheet 7b: Корзины по дням (добавления в корзину)
    add_dynamics_sheet(wb, "Корзины 14д", "8E44AD", pv_carts, "Корзины", add_period_sum=True)

    # Sheet 7c: CR клик → корзина по дням. Это средняя ступень воронки
    # (показ → КЛИК → КОРЗИНА → заказ): показывает, насколько карточка убеждает
    # уже зашедшего человека. Лист «Конверсия карточек» считает то же самое, но
    # СУММАРНО за период — здесь видно, в какой день просело.
    add_dynamics_sheet(wb, "CR клик→корзина 14д", "9B59B6", None, "CR%",
                       is_ratio=True, numerator_pv=pv_carts, denominator_pv=pv_clicks)

    # Sheet 8: Покупки (заказы)
    add_dynamics_sheet(wb, "Покупки 14д", "E67E22", pv_orders, "Покупки", add_period_sum=True)

    # Sheet 9: CR клик → покупка
    add_dynamics_sheet(wb, "CR клик→покупка 14д", "C0392B", None, "CR%",
                       is_ratio=True, numerator_pv=pv_orders, denominator_pv=pv_clicks)

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 9b: ВИДИМОСТЬ ↔ ДОЛЯ БУСТА ↔ ПОКАЗЫ (по дням, блоками по SKU)
    # ══════════════════════════════════════════════════════════════════════════
    if all_dates and active_skus:
        ws_vb = wb.create_sheet("Видимость↔Буст↔Показы")
        ws_vb.sheet_properties.tabColor = "16A085"

        N = len(all_dates)
        DAY_START = 4
        DAY_END = 3 + N
        COL_AVG = DAY_END + 1
        COL_RB = COL_AVG + 1   # corr boost-share → shows
        COL_RV = COL_RB + 1    # corr visibility → shows
        n_cols = COL_RV

        # Title
        ws_vb.merge_cells(f"A1:{get_column_letter(n_cols)}1")
        ws_vb["A1"] = ("ВИДИМОСТЬ ↔ ДОЛЯ БУСТА ↔ ПОКАЗЫ по дням  ·  "
                       f"{all_dates[0].strftime('%d.%m')}–{all_dates[-1].strftime('%d.%m.%Y')}")
        ws_vb["A1"].font = Font(bold=True, color="FFFFFF", size=14)
        ws_vb["A1"].fill = _fill("16A085")
        ws_vb["A1"].alignment = Alignment(horizontal="center")
        ws_vb.row_dimensions[1].height = 32

        # Header
        hdr_row = 2
        vb_heads = ["Товар", "SKU", "Метрика"] + [d.strftime("%d.%m") for d in all_dates] + \
                   ["Сред.", "r:буст→\nпоказы", "r:видим→\nпоказы"]
        for c, h in enumerate(vb_heads, 1):
            ws_vb.cell(row=hdr_row, column=c, value=h)
        _style_header_row(ws_vb, hdr_row, n_cols, "16A085")
        ws_vb.row_dimensions[hdr_row].height = 34

        def _corr(xs, ys):
            if len(xs) < 3:
                return None
            a, b = np.array(xs, dtype=float), np.array(ys, dtype=float)
            if np.std(a) == 0 or np.std(b) == 0:
                return None
            return round(float(np.corrcoef(a, b)[0, 1]), 2)

        # Per-SKU numeric series from df_full
        for sku in sorted(active_skus):
            name = name_map_full.get(sku, sku)
            sub = df_full[df_full["shopSku"] == sku] if "shopSku" in df_full.columns else df_full.iloc[0:0]
            by_date = {}
            for _, rr in sub.iterrows():
                by_date[rr["date"]] = rr

            shows_series, vis_series, bsh_series = [], [], []
            boost_series, org_series = [], []   # показы по бусту / органике (шт)
            for d in all_dates:
                rr = by_date.get(d)
                if rr is None:
                    shows_series.append(0); vis_series.append(None); bsh_series.append(None)
                    boost_series.append(0); org_series.append(0)
                    continue
                sh = pd.to_numeric(rr.get("shows"), errors="coerce")
                sb = pd.to_numeric(rr.get("shows_promo"), errors="coerce")
                vi = pd.to_numeric(rr.get("visibility"), errors="coerce")
                sh = 0 if pd.isna(sh) else int(sh)
                sbv = 0 if pd.isna(sb) else int(sb)
                shows_series.append(sh)
                boost_series.append(sbv)
                org_series.append(max(sh - sbv, 0))
                vis_series.append(None if pd.isna(vi) else round(float(vi), 1))
                if sh > 0:
                    bsh_series.append(round(sbv / sh * 100, 1))
                else:
                    bsh_series.append(None)

            # Correlations (pair only days where the metric is present)
            rb_pairs = [(b, s) for b, s in zip(bsh_series, shows_series) if b is not None]
            rv_pairs = [(v, s) for v, s in zip(vis_series, shows_series) if v is not None]
            r_boost = _corr([p[0] for p in rb_pairs], [p[1] for p in rb_pairs])
            r_vis = _corr([p[0] for p in rv_pairs], [p[1] for p in rv_pairs])

            base = ws_vb.max_row + 1
            shows_avg = round(sum(shows_series) / N, 1) if N else 0
            boost_avg = round(sum(boost_series) / N, 1) if N else 0
            org_avg = round(sum(org_series) / N, 1) if N else 0
            vis_nn = [v for v in vis_series if v is not None]
            bsh_nn = [b for b in bsh_series if b is not None]
            vis_avg = round(sum(vis_nn) / len(vis_nn), 1) if vis_nn else None
            bsh_avg = round(sum(bsh_nn) / len(bsh_nn), 1) if bsh_nn else None

            # Подписи строк счётных показов — для подсветки лейбла и формата
            COUNT_LABELS = ("Показы всего", "Показы орг.", "Показы буст")
            metric_rows = [
                ("Показы всего", shows_series, shows_avg, None),
                ("Показы орг.", org_series, org_avg, None),
                ("Показы буст", boost_series, boost_avg, None),
                ("Видим.%", vis_series, vis_avg, "0"),
                ("% буста", bsh_series, bsh_avg, "0.0"),
            ]
            BLOCK_ROWS = len(metric_rows)
            for i, (mlabel, series, avg, _fmt) in enumerate(metric_rows):
                r = base + i
                lcell = ws_vb.cell(row=r, column=3, value=mlabel)
                lcell.font = _tfont(bold=True)
                # тинт лейбла: органика — зелёный, буст — оранжевый
                if mlabel == "Показы орг.":
                    lcell.fill = _fill(GREEN_LT)
                elif mlabel == "Показы буст":
                    lcell.fill = _fill(ORANGE_LT)
                for j, d in enumerate(all_dates):
                    val = series[j]
                    cell = ws_vb.cell(row=r, column=DAY_START + j,
                                      value=(val if val is not None else "—"))
                    cell.border = _border()
                    cell.alignment = _center()
                    if mlabel in COUNT_LABELS:
                        cell.number_format = '#,##0'
                        if avg > 0 and val:
                            if val > avg * 1.3:
                                cell.fill = _fill(GREEN_LT)
                            elif val == 0:
                                cell.fill = _fill(RED_LIGHT)
                            elif val < avg * 0.5:
                                cell.fill = _fill(YELLOW_LT)
                    elif mlabel == "Видим.%":
                        cell.number_format = '0"%"'
                        if val is not None:
                            if val >= 70:
                                cell.fill = _fill(GREEN_LT)
                            elif val >= 40:
                                cell.fill = _fill(YELLOW_LT)
                            else:
                                cell.fill = _fill(RED_LIGHT)
                    else:  # % буста
                        cell.number_format = '0.0"%"'
                        if val is not None and val > 0:
                            cell.fill = _fill(ORANGE_LT if val >= 50 else YELLOW_LT)
                # Average column
                acell = ws_vb.cell(row=r, column=COL_AVG,
                                   value=(avg if avg is not None else "—"))
                acell.font = _tfont(bold=True)
                acell.border = _border()
                acell.alignment = _center()
                if mlabel in COUNT_LABELS:
                    acell.number_format = '#,##0'
                elif mlabel == "Видим.%":
                    acell.number_format = '0"%"'
                else:
                    acell.number_format = '0.0"%"'
                # Metric label cell border
                ws_vb.cell(row=r, column=3).border = _border()
                ws_vb.cell(row=r, column=3).alignment = _center()

            # Merge name / SKU / correlation columns across the metric-row block
            _bend = base + BLOCK_ROWS - 1
            ws_vb.merge_cells(start_row=base, start_column=1, end_row=_bend, end_column=1)
            ws_vb.merge_cells(start_row=base, start_column=2, end_row=_bend, end_column=2)
            ws_vb.merge_cells(start_row=base, start_column=COL_RB, end_row=_bend, end_column=COL_RB)
            ws_vb.merge_cells(start_row=base, start_column=COL_RV, end_row=_bend, end_column=COL_RV)

            ncell = ws_vb.cell(row=base, column=1, value=name[:45])
            ncell.alignment = Alignment(vertical="center", wrap_text=True)
            scell = ws_vb.cell(row=base, column=2, value=sku)
            scell.alignment = Alignment(vertical="center", horizontal="center")

            for cc, rval in ((COL_RB, r_boost), (COL_RV, r_vis)):
                rc = ws_vb.cell(row=base, column=cc,
                                value=(rval if rval is not None else "—"))
                rc.alignment = Alignment(vertical="center", horizontal="center")
                rc.font = _tfont(bold=True)
                if rval is not None:
                    if rval >= 0.5:
                        rc.fill = _fill(GREEN_LT); rc.font = _tfont(bold=True, color="1B5E20")
                    elif rval <= -0.5:
                        rc.fill = _fill(RED_LIGHT); rc.font = _tfont(bold=True, color="B71C1C")
                    else:
                        rc.fill = _fill(GREY_LT)

            # Borders around merged name/sku/corr cells
            for rr2 in range(base, base + BLOCK_ROWS):
                for cc2 in (1, 2, COL_RB, COL_RV):
                    ws_vb.cell(row=rr2, column=cc2).border = _border()

        # Column widths
        ws_vb.column_dimensions["A"].width = 34
        ws_vb.column_dimensions["B"].width = 14
        ws_vb.column_dimensions["C"].width = 10
        for i in range(DAY_START, DAY_END + 1):
            ws_vb.column_dimensions[get_column_letter(i)].width = 6.5
        ws_vb.column_dimensions[get_column_letter(COL_AVG)].width = 8
        ws_vb.column_dimensions[get_column_letter(COL_RB)].width = 9
        ws_vb.column_dimensions[get_column_letter(COL_RV)].width = 9
        ws_vb.freeze_panes = "D3"

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 10: ЮНИТ-ЭКОНОМИКА
    # ══════════════════════════════════════════════════════════════════════════
    all_fee_orders = kwargs.get("all_fee_orders")
    catalog_prices = kwargs.get("catalog_prices", {})
    promo_prices = kwargs.get("promo_prices", {})
    price_reco_ue = kwargs.get("price_reco", {}) or {}
    fee_detail = kwargs.get("fee_detail", [])
    ue_date_from = kwargs.get("date_from")
    ue_date_to = kwargs.get("date_to")
    boost_bids_ue = kwargs.get("boost_bids", {}) or {}
    boost_detail_ue = kwargs.get("boost_detail", {}) or {}
    netting_ue = kwargs.get("netting", {}) or {}
    if all_fee_orders:
        from .cost_prices import get_cost_price, TAX_RATE
        from .fbs_fby_client import DELIVERED_STATUSES, RETURN_STATUSES
        from collections import defaultdict

        ws10 = wb.create_sheet("Юнит-экономика")
        ws10.sheet_properties.tabColor = "16A085"

        # Period header
        period_str = ""
        if ue_date_from and ue_date_to:
            period_str = f" · {ue_date_from.strftime('%d.%m.%Y')} — {ue_date_to.strftime('%d.%m.%Y')}"
        num_ue_cols = 19  # will be set after headers defined
        ws10.merge_cells(f"A1:{get_column_letter(num_ue_cols)}1")
        ws10["A1"] = f"ЮНИТ-ЭКОНОМИКА{period_str}"
        ws10["A1"].font = Font(bold=True, color="FFFFFF", size=14)
        ws10["A1"].fill = _fill("16A085")
        ws10["A1"].alignment = Alignment(horizontal="center")
        ws10.row_dimensions[1].height = 36

        COMM_LABELS = [
            ("FEE",                    "Витрина"),
            ("AUCTION_PROMOTION",      "Буст"),
            ("DELIVERY_TO_CUSTOMER",   "Доставка\nпокуп."),
            ("CROSSREGIONAL_DELIVERY", "Средняя\nмиля"),
            ("PAYMENT_TRANSFER",       "Приём\nплатежа"),
            ("AGENCY",                 "Агент."),
            # SORTING / FULFILLMENT / RETURN_PROCESSING исключены: эти комиссии
            # начисляются только на RETURNED/CANCELLED заказах (идут в "Затраты
            # возвр."), на DELIVERED всегда 0 — для расчёта прибыли/ед. бесполезны.
        ]

        # Column layout:
        # A=Товар B=SKU C=FBS/FBY D=Доставлено E=Возвратов
        # F=Цена продажи (каталог) G=Факт. выручка/ед H=Софин. Яндекса I=Себест. J=Налог 7%
        # K.. = commission types (DELIVERED only, per unit)
        # далее: Итого комиссии, Чистая прибыль/ед (от факт.выручки), Промо цена,
        # Чистая прибыль (промо), Цена vs индекс, Маржа %, Затраты возвр.,
        # Переплата FEE (аномалии), Чистая выручка, Канал, Почему
        n_comm = len(COMM_LABELS)
        COL_DELIV_Y = 6     # F — доставлено вчера (динамика)
        COL_DELIV_DB = 7    # G — доставлено позавчера (динамика)
        COL_PRICE = 8       # H — каталожная цена (база для налога и комиссий)
        COL_REAL_REV = 9    # I — фактическая выручка/ед (цена покупателя + софин.)
        COL_COFIN = 10      # J — софинансирование Маркета (средн. ₽/ед., для справки)
        COL_COST = 11       # K
        COL_TAX = 12        # L
        COL_COMM_START = COL_TAX + 1
        COL_COMM_END = COL_COMM_START + n_comm - 1
        COL_TOTAL_COMM = COL_COMM_END + 1
        COL_PROFIT = COL_TOTAL_COMM + 1
        COL_PROFIT_FACT = COL_PROFIT + 1    # чистая прибыль/ед ФАКТ (из выплат netting)
        COL_PROFIT_FACT_TOT = COL_PROFIT_FACT + 1  # чистая прибыль ФАКТ ИТОГО по артикулу
        COL_PROMO_PRICE = COL_PROFIT_FACT_TOT + 1
        COL_PROMO_PROFIT = COL_PROMO_PRICE + 1
        COL_PROMO_INDEX = COL_PROMO_PROFIT + 1  # промо-цена vs индекс цен (выгодная?)
        COL_MARGIN = COL_PROMO_INDEX + 1        # маржа РАСЧЁТ (модель)
        COL_MARGIN_FACT = COL_MARGIN + 1        # маржа ФАКТ (из выплат)
        COL_RETURN_COST = COL_MARGIN_FACT + 1
        COL_FEE_OVERPAY = COL_RETURN_COST + 1   # переплата FEE из-за несработавшей скидки
        COL_NET_REVENUE = COL_FEE_OVERPAY + 1
        COL_BOOST_BID = COL_NET_REVENUE + 1      # текущая ставка буста % (снимок/API)
        COL_BOOST_MAXBID = COL_BOOST_BID + 1     # макс. ставка % для прибыли ≥ +100₽
        COL_BOOST_RATE = COL_BOOST_MAXBID + 1    # сработавшая ставка % (факт, UMS 30д)
        # ВАЖНО: NET — это реально списанные деньги (постоплата минус скидка за
        # совместные акции), именно он и есть «сколько стоил буст». GROSS — сработавший
        # аукцион ДО скидки; Яндекс срезает её сильно (на SKUfuk: gross 149 ₽/шт →
        # net 21 ₽/шт, т.е. −86%). Раньше в колонке «Буст факт» стоял GROSS, и расход
        # читался завышенным в разы. Держим оба: net как основной, gross — рядом.
        COL_BOOST_FACT = COL_BOOST_RATE + 1      # буст NET ₽/ед (реальное списание)
        COL_BOOST_GROSS = COL_BOOST_FACT + 1     # буст GROSS ₽/ед (до скидки акций)
        COL_REC = COL_BOOST_GROSS + 1       # рекомендация канала
        COL_REC_WHY = COL_REC + 1           # причина одной строкой
        num_ue_cols = COL_REC_WHY
        BOOST_FLOOR_PROFIT = 100  # минимальная прибыль/ед после буста (₽)

        # Fix merge for period header
        ws10.unmerge_cells(f"A1:{get_column_letter(19)}1")
        ws10.merge_cells(f"A1:{get_column_letter(num_ue_cols)}1")

        ue_heads = [
            "Товар", "SKU", "FBS/FBY", "Доставл.\nшт.", "Возвр.\nшт.",
            "Доставл.\nвчера", "Доставл.\nпозавч.",
            "Цена\nпродажи", "Факт.\nвыручка\n/ед.", "Софин.\nЯндекса\n/ед.", "Себест.", "Налог\n7%",
        ] + [lbl for _, lbl in COMM_LABELS] + [
            "Итого\nкомиссии", "Чистая\nприбыль\n/ед.", "Чистая\nприбыль/ед\nФАКТ",
            "Чистая\nприбыль ФАКТ\nитого ₽",
            "Промо\nцена", "Прибыль по\nпромо-цене\n(БЕЗ баллов)", "Цена vs\nиндекс цен",
            "Маржа %\n(расчёт)", "Маржа %\nФАКТ", "Затраты\nвозвр.",
            "Переплата\nFEE\n(аномалии)", "Чистая\nвыручка\n(расчёт)",
            "Ставка\nбуста %\n(тек)", "Макс.\nставка %\n(+100₽)",
            "Сраб.\nставка %\n(факт 14д)",
            "Буст NET\n₽/ед (14д)\nреально", "Буст GROSS\n₽/ед (14д)\nдо скидки",
            "Канал\n(реком.)", "Почему",
        ]

        hdr_row = 2
        for c, h in enumerate(ue_heads, 1):
            ws10.cell(row=hdr_row, column=c, value=h)
        _style_header_row(ws10, hdr_row, len(ue_heads), "16A085")
        ws10.row_dimensions[hdr_row].height = 40

        # Примечания к буст-колонкам: разница NET/GROSS неочевидна и раньше приводила
        # к чтению расхода завышенным в разы. Кладём в комментарий к заголовку, чтобы
        # не занимать строку под шапкой.
        ws10.cell(row=hdr_row, column=COL_BOOST_FACT).comment = Comment(
            "Буст NET — РЕАЛЬНО списанные деньги за буст на единицу:\n"
            "постоплата аукциона МИНУС скидка за участие в совместных акциях.\n"
            "Это и есть «сколько стоил буст». Ориентируйся на эту колонку.", "агент")
        ws10.cell(row=hdr_row, column=COL_BOOST_GROSS).comment = Comment(
            "Буст GROSS — сработавшая ставка аукциона ДО скидки за совместные акции.\n"
            "НЕ расход: Яндекс срезает скидкой обычно большую часть.\n"
            "Пример (SKUfuk, 14д, 8 продаж): gross 149,2 ₽/шт → net 20,6 ₽/шт (−86%).\n"
            "Полезен, чтобы видеть, как отработал аукцион (сраб. ставка обычно ниже "
            "выставленной — аукцион второй цены).", "агент")
        ws10.cell(row=hdr_row, column=COL_PROMO_PROFIT).comment = Comment(
            "МОДЕЛЬ, а не факт: промо-цена − налог 7% − комиссии − себест.\n"
            "Баллы Яндекса в формулу НЕ входят, поэтому промо-прибыль занижена:\n"
            "по части скидок Яндекс доплачивает баллами («Баллы за скидку Маркета»).\n"
            "ФАКТ по деньгам — лист «Сверка выплат»: снижение цены там уже сидит\n"
            "в «Платежах покуп.» (покупатель просто заплатил меньше), а колонки\n"
            "🟢 «БАЛЛЫ ₽/шт» и 🔴 «УСЛУГИ ЗА СЧЁТ БАЛЛОВ ₽/шт» показывают замкнутый\n"
            "контур бонусного баланса: Яндекс кладёт баллы за скидку, которую\n"
            "профинансировал сам, и тут же списывает ими твои услуги (размещение/\n"
            "буст/логистика). Разница осмысленна по каталогу (+6–10%), но не по SKU.", "агент")

        # Aggregate per (sku, type).
        # Commissions split: delivered vs returned, so return costs
        # don't inflate per-unit metrics for delivered items.
        agg = defaultdict(lambda: {
            "name": "", "delivered_count": 0, "return_count": 0,
            "delivered_y": 0, "delivered_db": 0,   # доставлено вчера / позавчера (динамика)
            "delivered_comm": defaultdict(float),
            "return_comm_total": 0.0,
            "delivered_cofin": 0.0,    # Σ софинансирования Маркета по доставленным
            "delivered_revenue": 0.0,  # Σ фактической выручки (покупатель + софин.)
            "fee_overpay": 0.0,        # Σ FEE на заказах, где скидка за акцию не сработала
        })
        _y_iso = yesterday.isoformat()
        _db_iso = day_before.isoformat()

        for label in ("fbs", "fby"):
            tag = label.upper()
            for o in all_fee_orders.get(label, []):
                key = (o["sku"], tag)
                a = agg[key]
                a["name"] = a["name"] or o["name"]
                cnt = o["count"]
                is_delivered = o["status"] in DELIVERED_STATUSES

                if is_delivered:
                    a["delivered_count"] += cnt
                    sud = o.get("status_update_date", "")
                    if sud == _y_iso:
                        a["delivered_y"] += cnt
                    elif sud == _db_iso:
                        a["delivered_db"] += cnt
                    buyer = o.get("buyer_price", 0)
                    mp = o.get("marketplace_price", 0)
                    # Софинансирование Маркета (доплата сверх цены покупателя)
                    a["delivered_cofin"] += mp * cnt
                    # Фактическая выручка продавца = цена покупателя + софинансирование
                    a["delivered_revenue"] += (buyer + mp) * cnt
                    # Аномалия FEE: скидка за акцию не сработала → FEE = полный тариф
                    # (норм. FEE после скидки ~0; берём порог 10% от номинала)
                    fee_pu = o["commissions"].get("FEE", 0)
                    nominal_pu = buyer + mp
                    if nominal_pu > 0 and fee_pu > nominal_pu * 0.10:
                        a["fee_overpay"] += fee_pu * cnt
                    for ctype, cval in o["commissions"].items():
                        a["delivered_comm"][ctype] += cval * cnt
                else:
                    a["return_count"] += cnt
                    for cval in o["commissions"].values():
                        a["return_comm_total"] += cval * cnt

        # Sort by SKU name
        sorted_keys = sorted(agg.keys(), key=lambda k: k[0])

        # ── Рекомендация канала FBS/FBY/оба (на уровне товара) ──
        # Основа — скорость продаж (гладкая метрика), маржа — вторичный флаг,
        # сток FBY — предохранитель. См. логику в _channel_recommendation ниже.
        fee_window_days = kwargs.get("fee_window_days") or 20
        weeks = max(fee_window_days / 7.0, 1.0)
        fby_stocks_ue = stocks.get("fby", {})

        # Соберём по товару: продажи в каждом канале + маржа в каждом канале
        per_sku = defaultdict(lambda: {"FBS": {"deliv": 0, "margin": None},
                                        "FBY": {"deliv": 0, "margin": None}})
        for (sku, tag), a in agg.items():
            d = a["delivered_count"]
            per_sku[sku][tag]["deliv"] = d
            price = catalog_prices.get(sku, 0)
            cost = get_cost_price(sku) or 0
            if d > 0 and price > 0:
                totcomm = sum(a["delivered_comm"].values()) / d
                profit = price - round(price * 0.07, 2) - totcomm - cost
                per_sku[sku][tag]["margin"] = profit / price * 100

        def _channel_recommendation(sku):
            """Returns (label, why) — товар-уровневая рекомендация."""
            ps = per_sku[sku]
            s_d, y_d = ps["FBS"]["deliv"], ps["FBY"]["deliv"]
            total = s_d + y_d
            per_week = total / weeks
            fby_avail = fby_stocks_ue.get(sku, {}).get("available", 0)

            # Мало данных — не гадаем
            if total < 4:
                return "мало данных", f"всего {total} продаж за {fee_window_days}д"

            # Вторичный маржа-флаг: FBS заметно выгоднее по марже
            margin_flag = ""
            s_m, y_m = ps["FBS"]["margin"], ps["FBY"]["margin"]
            if s_m is not None and y_m is not None and (s_m - y_m) > 8:
                margin_flag = f"; ⚠ FBS маржа +{round(s_m - y_m)}пп"

            # Базовое решение по скорости
            if per_week >= 5:
                base = "FBY"
                why = f"{per_week:.1f} шт/нед — быстрый, буст+витрина выгоднее"
                if fby_avail == 0:
                    return "FBY — завезти", f"{why}; сток FBY = 0" + margin_flag
                if margin_flag:
                    return "FBY*", why + margin_flag
                return base, why
            elif per_week >= 2:
                why = f"{per_week:.1f} шт/нед — средний оборот"
                return "оба", why + margin_flag
            else:
                return "FBS", f"{per_week:.1f} шт/нед — медленный, не морозить сток FBY" + margin_flag

        missing_costs = []
        data_start_row = hdr_row + 1
        ue_summary = []  # for ТОП-10 по чистой выручке: [(name, sku, tag, delivered, net_est)]

        for key in sorted_keys:
            sku, tag = key
            a = agg[key]
            delivered = a["delivered_count"]
            returns = a["return_count"]
            total_items = delivered + returns
            if total_items == 0:
                continue

            cat_price = catalog_prices.get(sku, 0)
            cost = get_cost_price(sku)
            if cost is None:
                missing_costs.append(sku)
                cost = 0

            # Per-unit commission averages from DELIVERED orders only
            comm_vals = []
            if delivered > 0:
                for ctype, _ in COMM_LABELS:
                    comm_vals.append(round(a["delivered_comm"].get(ctype, 0) / delivered, 2))
            else:
                comm_vals = [0.0] * n_comm

            return_cost_total = round(a["return_comm_total"], 2)

            r = ws10.max_row + 1
            ws10.cell(row=r, column=1, value=a["name"][:50])
            ws10.cell(row=r, column=2, value=sku)
            ws10.cell(row=r, column=3, value=tag)
            ws10.cell(row=r, column=4, value=delivered)
            ws10.cell(row=r, column=5, value=returns)
            # Динамика доставок: вчера / позавчера
            dy = a["delivered_y"]; ddb = a["delivered_db"]
            cy = ws10.cell(row=r, column=COL_DELIV_Y, value=dy)
            cdb = ws10.cell(row=r, column=COL_DELIV_DB, value=ddb)
            if dy > ddb:
                cy.font = _tfont(bold=True, color="1B5E20")
            elif dy < ddb:
                cy.font = _tfont(bold=True, color="B71C1C")
            ws10.cell(row=r, column=COL_PRICE, value=cat_price)
            # Фактическая выручка/ед = (цена покупателя + софин.) / доставлено.
            # Если данных о ценах нет — откатываемся на каталожную цену.
            real_rev = (round(a["delivered_revenue"] / delivered, 2)
                        if delivered > 0 and a["delivered_revenue"] > 0 else cat_price)
            ws10.cell(row=r, column=COL_REAL_REV, value=real_rev)
            # Софинансирование Маркета: средняя доплата Яндекса на единицу (для справки)
            cofin = round(a["delivered_cofin"] / delivered, 2) if delivered > 0 else 0
            ws10.cell(row=r, column=COL_COFIN, value=cofin)
            ws10.cell(row=r, column=COL_COST, value=cost)

            # Tax = 7% of price (Excel formula)
            price_ref = f"{get_column_letter(COL_PRICE)}{r}"
            cost_ref = f"{get_column_letter(COL_COST)}{r}"
            tax_ref = f"{get_column_letter(COL_TAX)}{r}"
            ws10.cell(row=r, column=COL_TAX, value=f"=ROUND({price_ref}*0.07,2)")

            # Commission values (delivered only, per unit)
            for i, val in enumerate(comm_vals):
                ws10.cell(row=r, column=COL_COMM_START + i, value=val)

            # Total commissions (formula: sum of commission columns)
            comm_start_letter = get_column_letter(COL_COMM_START)
            comm_end_letter = get_column_letter(COL_COMM_END)
            ws10.cell(row=r, column=COL_TOTAL_COMM,
                value=f"=SUM({comm_start_letter}{r}:{comm_end_letter}{r})")

            # Прибыль/ед = Факт.выручка − Налог(от каталожной) − Комиссии(факт) − Себест.
            # Выручка берётся фактическая (с софин.), а налог 7% и комиссии — от
            # каталожной/номинальной базы, как их начисляет Яндекс.
            real_rev_ref = f"{get_column_letter(COL_REAL_REV)}{r}"
            total_comm_ref = f"{get_column_letter(COL_TOTAL_COMM)}{r}"
            delivered_ref = f"{get_column_letter(4)}{r}"
            ws10.cell(row=r, column=COL_PROFIT,
                value=f"=IF({delivered_ref}>0,{real_rev_ref}-{tax_ref}-{total_comm_ref}-{cost_ref},0)")

            # Margin % = Profit / Факт.выручка (formula)
            profit_ref = f"{get_column_letter(COL_PROFIT)}{r}"
            ws10.cell(row=r, column=COL_MARGIN,
                value=f"=IF(AND({real_rev_ref}>0,{delivered_ref}>0),ROUND({profit_ref}/{real_rev_ref}*100,1),0)")

            # Promo price + promo profit (same formula, with promo price instead of catalog price)
            promo = promo_prices.get(sku)
            if promo:
                promo_price = promo["promo_price"]
                ws10.cell(row=r, column=COL_PROMO_PRICE, value=promo_price)
                promo_price_ref = f"{get_column_letter(COL_PROMO_PRICE)}{r}"
                ws10.cell(row=r, column=COL_PROMO_PROFIT,
                    value=f"=IF({delivered_ref}>0,{promo_price_ref}-ROUND({promo_price_ref}*0.07,2)"
                          f"-{total_comm_ref}-{cost_ref},0)")
                compare_price = promo_price
            else:
                compare_price = cat_price

            # Цена (промо, если есть, иначе обычная цена продажи) vs индекс цен:
            # стала ли выгодной / сколько ещё снизить
            if compare_price:
                reco = price_reco_ue.get(sku) or {}
                optimal = reco.get("optimal_price")
                pic = ws10.cell(row=r, column=COL_PROMO_INDEX)
                if optimal is None:
                    pic.value = "нет индекса"
                    pic.font = _tfont(color="999999")
                elif compare_price <= optimal:
                    pic.value = f"выгодная ✓ (порог {optimal:.0f}₽)"
                    pic.fill = _fill(GREEN_LT); pic.font = _tfont(bold=True, color="1B5E20")
                else:
                    gap = compare_price - optimal
                    pic.value = f"−{gap:.0f}₽ до выгодной (порог {optimal:.0f}₽)"
                    pic.fill = _fill(YELLOW_LT if gap <= optimal * 0.1 else RED_LIGHT)
                    pic.font = _tfont(bold=True,
                                      color="9A7D0A" if gap <= optimal * 0.1 else "B71C1C")

            # Return costs (total, not per-unit)
            ws10.cell(row=r, column=COL_RETURN_COST, value=return_cost_total)

            # Переплата FEE: суммарный FEE на заказах, где скидка за акцию не сработала
            fee_overpay = round(a["fee_overpay"], 2)
            ws10.cell(row=r, column=COL_FEE_OVERPAY, value=fee_overpay)

            # Net revenue = Profit/unit * Delivered - ReturnCosts (formula)
            delivered_ref = f"{get_column_letter(4)}{r}"
            return_cost_ref = f"{get_column_letter(COL_RETURN_COST)}{r}"
            ws10.cell(row=r, column=COL_NET_REVENUE,
                value=f"={profit_ref}*{delivered_ref}-{return_cost_ref}")

            # Рекомендация канала (одинаковая для FBS/FBY строк одного товара)
            rec_label, rec_why = _channel_recommendation(sku)
            rc = ws10.cell(row=r, column=COL_REC, value=rec_label)
            ws10.cell(row=r, column=COL_REC_WHY, value=rec_why)
            if rec_label.startswith("FBY"):
                rc.fill = _fill("D5F5E3" if "завезти" not in rec_label else ORANGE_LT)
                rc.font = _tfont(bold=True, color="1B5E20" if "завезти" not in rec_label else "E65100")
            elif rec_label == "FBS":
                rc.fill = _fill("D6EAF8")
                rc.font = _tfont(bold=True, color="1F618D")
            elif rec_label == "оба":
                rc.fill = _fill(YELLOW_LT)
                rc.font = _tfont(bold=True, color="9A7D0A")
            else:  # мало данных
                rc.font = _tfont(color="999999")

            # Styling
            for c in range(1, num_ue_cols + 1):
                cell = ws10.cell(row=r, column=c)
                cell.border = _border()
                cell.alignment = (_center() if c > 2 else _left()) if c != COL_REC_WHY else _left()
                if COL_PRICE <= c <= COL_COMM_END or c in (
                    COL_TOTAL_COMM, COL_PROFIT, COL_PROMO_PRICE, COL_PROMO_PROFIT,
                    COL_RETURN_COST, COL_FEE_OVERPAY, COL_NET_REVENUE):
                    cell.number_format = '#,##0.00'
                if c == COL_MARGIN:
                    cell.number_format = '0.0"%"'
                if c in (COL_BOOST_BID, COL_BOOST_MAXBID, COL_BOOST_RATE) and isinstance(cell.value, (int, float)):
                    cell.number_format = '0.0"%"'
                if c in (COL_BOOST_FACT, COL_BOOST_GROSS) and isinstance(cell.value, (int, float)):
                    cell.number_format = '#,##0.00'

            # Софинансирование: голубой фон; жирно-синий, если доля заметная
            # (Яндекс доплачивает >10% от каталожной цены)
            if delivered > 0 and cofin > 0:
                cfcell = ws10.cell(row=r, column=COL_COFIN)
                cfcell.fill = _fill(BLUE_LT)
                if cat_price > 0 and cofin > cat_price * 0.1:
                    cfcell.font = _tfont(bold=True, color="1F618D")

            # Color profit cell (выручка факт., налог/комиссии от каталожной)
            total_fees_per_unit = sum(comm_vals)
            tax_est = cat_price * 0.07
            profit_est = (real_rev - tax_est - total_fees_per_unit - cost) if delivered > 0 else 0
            pcell = ws10.cell(row=r, column=COL_PROFIT)
            if profit_est > 0:
                pcell.fill = _fill(GREEN_LT)
                pcell.font = _tfont(bold=True, color="1B5E20")
            elif profit_est < 0:
                pcell.fill = _fill(RED_LIGHT)
                pcell.font = _tfont(bold=True, color="B71C1C")

            # ── Чистая прибыль/ед ФАКТ (из реальных выплат netting за 30д) ──
            # Сравниваем расчётную прибыль с фактической: видно, сколько на самом
            # деле компенсирует Яндекс (зелёным — факт ≥ расчёта, красным — меньше).
            nt = netting_ue.get((sku, tag))
            if nt and nt.get("qty", 0) > 0:
                net_q = nt["qty"] - nt.get("ret_qty", 0)   # чистые штуки = продано − возвраты
                # Возвраты годные → себест на чистые штуки, налог с нетто-выручки.
                fact_profit_tot = round(nt["payout"] - cost * net_q
                                        - (nt["buyer_pay"] + nt.get("returns", 0)) * 0.07, 2)
                fact_profit_pu = round(fact_profit_tot / net_q, 2) if net_q > 0 else 0
                fact_margin = round(fact_profit_tot / nt["buyer_pay"] * 100, 1) if nt["buyer_pay"] else 0
                fcell = ws10.cell(row=r, column=COL_PROFIT_FACT, value=fact_profit_pu)
                fcell.number_format = '#,##0.00'
                ftot = ws10.cell(row=r, column=COL_PROFIT_FACT_TOT, value=fact_profit_tot)
                ftot.number_format = '#,##0'
                fmar = ws10.cell(row=r, column=COL_MARGIN_FACT, value=fact_margin)
                fmar.number_format = '0.0"%"'
                # подсветка: факт/ед vs расчёт/ед
                clr = "1B5E20" if fact_profit_pu >= profit_est else "B71C1C"
                fill = GREEN_LT if fact_profit_pu >= profit_est else RED_LIGHT
                for cc in (fcell, ftot):
                    cc.fill = _fill(fill); cc.font = _tfont(bold=True, color=clr)
                # маржа факт: своя подсветка по уровню
                fmar.fill = _fill(GREEN_LT if fact_margin >= 20 else RED_LIGHT if fact_margin < 5 else YELLOW_LT)
            else:
                ws10.cell(row=r, column=COL_PROFIT_FACT, value="—")
                ws10.cell(row=r, column=COL_PROFIT_FACT_TOT, value="—")
                ws10.cell(row=r, column=COL_MARGIN_FACT, value="—")

            # Color promo profit cell
            if promo:
                tax_promo_est = promo_price * 0.07
                promo_profit_est = (promo_price - tax_promo_est - total_fees_per_unit - cost) if delivered > 0 else 0
                ppcell = ws10.cell(row=r, column=COL_PROMO_PROFIT)
                if promo_profit_est > 0:
                    ppcell.fill = _fill(GREEN_LT)
                    ppcell.font = _tfont(bold=True, color="1B5E20")
                elif promo_profit_est < 0:
                    ppcell.fill = _fill(RED_LIGHT)
                    ppcell.font = _tfont(bold=True, color="B71C1C")

            # ── Экономика буста: предел ставки для прибыли ≥ +100₽/ед ──
            # Цена продажи = промо-цена (что платит покупатель), иначе факт.выручка.
            # Прибыль ДО буста = цена − налог7% − комиссии − себест (буст списывается
            # отдельно «оплатой за продажу», в комиссиях его нет).
            sale_price = promo_price if promo else real_rev
            profit_before_boost = (sale_price - sale_price * 0.07 - total_fees_per_unit - cost) \
                if (delivered > 0 and sale_price) else 0
            max_boost_spend = round(profit_before_boost - BOOST_FLOOR_PROFIT, 1)
            max_boost_bid = round(max_boost_spend / sale_price * 100, 1) if sale_price else None
            cur_bid = boost_bids_ue.get(sku, {}).get("bid_pct")
            ws10.cell(row=r, column=COL_BOOST_BID, value=cur_bid if cur_bid is not None else "—")
            mbcell = ws10.cell(row=r, column=COL_BOOST_MAXBID,
                               value=max_boost_bid if (max_boost_bid is not None and max_boost_bid > 0) else "—")
            # ФАКТ. буст из UMS «Буст продаж»: сработавшая ставка % + NET и GROSS ₽/ед.
            # NET = постоплата − скидка за совм. акции = реально ушедшие деньги.
            bd = boost_detail_ue.get(sku, {})
            rate_avg = bd.get("rate_avg")
            gross_pu = bd.get("gross_per_unit")
            net_pu = bd.get("net_per_unit")
            ws10.cell(row=r, column=COL_BOOST_RATE,
                      value=rate_avg if rate_avg else "—")
            nb = ws10.cell(row=r, column=COL_BOOST_FACT,
                           value=net_pu if net_pu else "—")
            gb = ws10.cell(row=r, column=COL_BOOST_GROSS,
                           value=gross_pu if gross_pu else "—")
            # NET — основная цифра, выделяем; GROSS приглушаем, чтобы его не читали
            # как расход (он до скидки и обычно в разы больше).
            if isinstance(net_pu, (int, float)) and net_pu:
                nb.font = _tfont(bold=True)
                # дорогой буст: > 100 ₽/ед — красный, 50–100 — жёлтый, дешевле — без фона
                if net_pu > 100:
                    nb.fill = _fill(RED_LIGHT); nb.font = _tfont(bold=True, color="B71C1C")
                elif net_pu >= 50:
                    nb.fill = _fill(YELLOW_LT); nb.font = _tfont(bold=True, color="7D6608")
            if isinstance(gross_pu, (int, float)) and gross_pu:
                gb.font = _tfont(color="888888")
            # подсветка: если макс.ставка отрицательна/мала — буст невыгоден
            if isinstance(max_boost_bid, (int, float)):
                if max_boost_bid <= 0:
                    mbcell.fill = _fill(RED_LIGHT); mbcell.font = _tfont(bold=True, color="B71C1C")
                elif isinstance(cur_bid, (int, float)) and cur_bid > max_boost_bid:
                    # текущая ставка выше потолка — риск уйти ниже +100₽
                    mbcell.fill = _fill(ORANGE_LT); mbcell.font = _tfont(bold=True, color="E65100")
                else:
                    mbcell.fill = _fill(GREEN_LT)

            # Color net revenue cell
            net_est = profit_est * delivered - return_cost_total
            ue_summary.append((a["name"][:50], sku, tag, delivered, net_est))
            nrcell = ws10.cell(row=r, column=COL_NET_REVENUE)
            if net_est > 0:
                nrcell.fill = _fill(GREEN_LT)
                nrcell.font = _tfont(bold=True, color="1B5E20")
            elif net_est < 0:
                nrcell.fill = _fill(RED_LIGHT)
                nrcell.font = _tfont(bold=True, color="B71C1C")

            # Color return cost cell if > 0
            if return_cost_total > 0:
                ws10.cell(row=r, column=COL_RETURN_COST).fill = _fill(ORANGE_LT)
                ws10.cell(row=r, column=COL_RETURN_COST).font = _tfont(bold=True, color="E65100")

            # Color FEE overpay cell if > 0 (красный — реальная утечка денег)
            if fee_overpay > 0:
                fo = ws10.cell(row=r, column=COL_FEE_OVERPAY)
                fo.fill = _fill(RED_LIGHT)
                fo.font = _tfont(bold=True, color="B71C1C")

            # Color cost cell if missing
            if get_cost_price(sku) is None:
                ws10.cell(row=r, column=COL_COST).fill = _fill(YELLOW_LT)
                ws10.cell(row=r, column=COL_COST).font = _tfont(bold=True, color="F57F17")

            # Color returns count cell if > 0
            if returns > 0:
                ws10.cell(row=r, column=5).fill = _fill(ORANGE_LT)
                ws10.cell(row=r, column=5).font = _tfont(bold=True, color="E65100")

        # Summary row with formulas
        ws10.append([])
        sr = ws10.max_row + 1
        last_data = sr - 2  # last data row (before empty row)
        ws10.cell(row=sr, column=1, value="ИТОГО").font = _tfont(sz=12, bold=True)

        sum_cols = [4, 5, COL_DELIV_Y, COL_DELIV_DB, COL_PRICE, COL_COST, COL_TAX] + \
                   list(range(COL_COMM_START, COL_COMM_END + 1)) + \
                   [COL_TOTAL_COMM, COL_PROFIT, COL_RETURN_COST, COL_FEE_OVERPAY, COL_NET_REVENUE]
        for col in sum_cols:
            letter = get_column_letter(col)
            ws10.cell(row=sr, column=col,
                value=f"=SUM({letter}{data_start_row}:{letter}{last_data})")
            ws10.cell(row=sr, column=col).font = _tfont(bold=True)
            ws10.cell(row=sr, column=col).number_format = '#,##0.00'

        # Margin % for totals
        ws10.cell(row=sr, column=COL_MARGIN,
            value=f"=IF({get_column_letter(COL_PRICE)}{sr}>0,"
                  f"ROUND({get_column_letter(COL_PROFIT)}{sr}"
                  f"/{get_column_letter(COL_PRICE)}{sr}*100,1),0)")
        ws10.cell(row=sr, column=COL_MARGIN).font = _tfont(bold=True)

        for c in range(1, num_ue_cols + 1):
            ws10.cell(row=sr, column=c).border = _border()
            ws10.cell(row=sr, column=c).alignment = _center() if c > 2 else _left()

        # Missing costs warning
        filter_last_row = sr

        if missing_costs:
            wr = sr + 2
            ws10.cell(row=wr, column=1,
                value=f"⚠ Себестоимость не найдена для: {', '.join(sorted(set(missing_costs)))}")
            ws10.cell(row=wr, column=1).font = _tfont(bold=True, color="B71C1C")
            ws10.merge_cells(start_row=wr, start_column=1, end_row=wr, end_column=num_ue_cols)

        # ── ТОП-10 ПО ЧИСТОЙ ВЫРУЧКЕ ──
        if ue_summary:
            top10 = sorted(ue_summary, key=lambda x: x[4], reverse=True)[:10]

            tr0 = ws10.max_row + 2
            ws10.merge_cells(start_row=tr0, start_column=1, end_row=tr0, end_column=5)
            ws10.cell(row=tr0, column=1, value="ТОП-10 ПО ЧИСТОЙ ВЫРУЧКЕ")
            ws10.cell(row=tr0, column=1).font = Font(bold=True, color="FFFFFF", size=12)
            ws10.cell(row=tr0, column=1).fill = _fill("16A085")
            ws10.cell(row=tr0, column=1).alignment = Alignment(horizontal="center")

            top_heads = ["Товар", "SKU", "FBS/FBY", "Доставл.\nшт.", "Чистая\nвыручка,₽"]
            thr = tr0 + 1
            for c, h in enumerate(top_heads, 1):
                ws10.cell(row=thr, column=c, value=h)
            _style_header_row(ws10, thr, len(top_heads), "16A085")

            for name, sku, tag, delivered, net_est in top10:
                r = ws10.max_row + 1
                ws10.cell(row=r, column=1, value=name)
                ws10.cell(row=r, column=2, value=sku)
                ws10.cell(row=r, column=3, value=tag)
                ws10.cell(row=r, column=4, value=delivered)
                nrc = ws10.cell(row=r, column=5, value=round(net_est, 2))
                nrc.number_format = '#,##0.00'
                if net_est > 0:
                    nrc.fill = _fill(GREEN_LT)
                    nrc.font = _tfont(bold=True, color="1B5E20")
                elif net_est < 0:
                    nrc.fill = _fill(RED_LIGHT)
                    nrc.font = _tfont(bold=True, color="B71C1C")
                for c in range(1, len(top_heads) + 1):
                    cell = ws10.cell(row=r, column=c)
                    cell.border = _border()
                    cell.alignment = _center() if c > 2 else _left()

        ws10.column_dimensions["A"].width = 42
        ws10.column_dimensions["B"].width = 16
        ws10.column_dimensions["C"].width = 10
        for i in range(4, num_ue_cols + 1):
            ws10.column_dimensions[get_column_letter(i)].width = 12
        ws10.column_dimensions[get_column_letter(COL_PROMO_INDEX)].width = 26
        ws10.column_dimensions[get_column_letter(COL_REC)].width = 14
        ws10.column_dimensions[get_column_letter(COL_REC_WHY)].width = 46
        ws10.freeze_panes = "D3"
        ws10.auto_filter.ref = f"A2:{get_column_letter(num_ue_cols)}{filter_last_row}"

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 11: ЛОГИСТИКА ПО РЕГИОНАМ
    # ══════════════════════════════════════════════════════════════════════════
    if all_fee_orders:
        from .fbs_fby_client import DELIVERED_STATUSES
        from collections import defaultdict

        ws11 = wb.create_sheet("Логистика по регионам")
        ws11.sheet_properties.tabColor = "D35400"

        period_str = ""
        if ue_date_from and ue_date_to:
            period_str = f" · {ue_date_from.strftime('%d.%m.%Y')} — {ue_date_to.strftime('%d.%m.%Y')}"

        # Collect all regions and per-SKU delivery costs
        # Key: (sku, region) → {delivery_sum, crossregional_sum, count}
        region_data = defaultdict(lambda: {
            "name": "", "delivery_sum": 0.0, "cross_sum": 0.0, "count": 0,
        })
        all_regions = set()

        for label in ("fbs", "fby"):
            for o in all_fee_orders.get(label, []):
                if o["status"] not in DELIVERED_STATUSES:
                    continue
                region = o.get("delivery_region", "") or "Неизвестно"
                sku = o["sku"]
                cnt = o["count"]
                # per-unit values from get_orders_with_fees
                delivery = o["commissions"].get("DELIVERY_TO_CUSTOMER", 0)
                cross = o["commissions"].get("CROSSREGIONAL_DELIVERY", 0)

                key = (sku, region)
                rd = region_data[key]
                rd["name"] = rd["name"] or o["name"]
                rd["delivery_sum"] += delivery * cnt
                rd["cross_sum"] += cross * cnt
                rd["count"] += cnt
                all_regions.add(region)

        # Also build per-SKU totals
        sku_totals = defaultdict(lambda: {
            "name": "", "delivery_sum": 0.0, "cross_sum": 0.0, "count": 0,
        })
        for (sku, region), rd in region_data.items():
            st = sku_totals[sku]
            st["name"] = st["name"] or rd["name"]
            st["delivery_sum"] += rd["delivery_sum"]
            st["cross_sum"] += rd["cross_sum"]
            st["count"] += rd["count"]

        sorted_regions = sorted(all_regions)

        # --- Sub-sheet A: Сводная по SKU с разбивкой по регионам ---
        # Header
        n_reg_cols = 4 + len(sorted_regions) * 2  # SKU, Товар, шт, ср.логистика, then per-region pairs
        ws11.merge_cells(f"A1:{get_column_letter(min(n_reg_cols, 16384))}1")
        ws11["A1"] = f"ЛОГИСТИКА ПО РЕГИОНАМ (доставка + средняя миля){period_str}"
        ws11["A1"].font = Font(bold=True, color="FFFFFF", size=14)
        ws11["A1"].fill = _fill("D35400")
        ws11["A1"].alignment = Alignment(horizontal="center")
        ws11.row_dimensions[1].height = 36

        # Build flat table: SKU | Товар | Шт | Ср.логистика | per-region columns
        hdr_row = 2
        reg_heads = ["SKU", "Товар", "Всего\nшт.", "Ср. логист.\nна ед."]
        for reg in sorted_regions:
            reg_short = reg[:15]
            reg_heads.append(f"{reg_short}\nшт.")
            reg_heads.append(f"{reg_short}\nлогист.")
        for c, h in enumerate(reg_heads, 1):
            ws11.cell(row=hdr_row, column=c, value=h)
        _style_header_row(ws11, hdr_row, len(reg_heads), "D35400")
        ws11.row_dimensions[hdr_row].height = 44

        data_start = hdr_row + 1
        for sku in sorted(sku_totals.keys()):
            st = sku_totals[sku]
            if st["count"] == 0:
                continue
            avg_logistics = round((st["delivery_sum"] + st["cross_sum"]) / st["count"], 2)

            r = ws11.max_row + 1
            ws11.cell(row=r, column=1, value=sku)
            ws11.cell(row=r, column=2, value=st["name"][:45])
            ws11.cell(row=r, column=3, value=st["count"])
            ws11.cell(row=r, column=4, value=avg_logistics)

            col = 5
            for reg in sorted_regions:
                rd = region_data.get((sku, reg))
                if rd and rd["count"] > 0:
                    avg_log = round((rd["delivery_sum"] + rd["cross_sum"]) / rd["count"], 2)
                    ws11.cell(row=r, column=col, value=rd["count"])
                    ws11.cell(row=r, column=col + 1, value=avg_log)
                col += 2

            # Styling
            for c in range(1, len(reg_heads) + 1):
                cell = ws11.cell(row=r, column=c)
                cell.border = _border()
                cell.alignment = _center() if c > 2 else _left()
                if c == 4 or (c >= 6 and c % 2 == 0):
                    cell.number_format = '#,##0.00'
                if c == 3 or (c >= 5 and c % 2 == 1):
                    cell.number_format = '#,##0'

            # Heatmap avg logistics
            avg_cell = ws11.cell(row=r, column=4)
            if avg_logistics > 100:
                avg_cell.fill = _fill(RED_LIGHT)
                avg_cell.font = _tfont(bold=True, color="B71C1C")
            elif avg_logistics > 50:
                avg_cell.fill = _fill(ORANGE_LT)
            elif avg_logistics > 0:
                avg_cell.fill = _fill(GREEN_LT)

        # Summary row: total units + avg logistics per region
        last_data_row = ws11.max_row
        ws11.append([])
        sr = ws11.max_row + 1
        ws11.cell(row=sr, column=1, value="ИТОГО").font = _tfont(sz=12, bold=True)

        # Col C: total units (SUM formula)
        c_letter = get_column_letter(3)
        ws11.cell(row=sr, column=3,
            value=f"=SUM({c_letter}{data_start}:{c_letter}{last_data_row})")

        # Col D: overall avg logistics (weighted)
        # = total_delivery_cost / total_units
        region_totals = defaultdict(lambda: {"count": 0, "cost_sum": 0.0})
        grand_count = 0
        grand_cost = 0.0
        for (sku, reg), rd in region_data.items():
            region_totals[reg]["count"] += rd["count"]
            region_totals[reg]["cost_sum"] += rd["delivery_sum"] + rd["cross_sum"]
            grand_count += rd["count"]
            grand_cost += rd["delivery_sum"] + rd["cross_sum"]

        ws11.cell(row=sr, column=4,
            value=round(grand_cost / grand_count, 2) if grand_count else 0)

        # Per-region columns: units sum + avg logistics
        col = 5
        for reg in sorted_regions:
            rt = region_totals.get(reg)
            if rt and rt["count"] > 0:
                ws11.cell(row=sr, column=col, value=rt["count"])
                ws11.cell(row=sr, column=col + 1,
                    value=round(rt["cost_sum"] / rt["count"], 2))
            col += 2

        # Style summary row
        for c in range(1, len(reg_heads) + 1):
            cell = ws11.cell(row=sr, column=c)
            cell.border = _border()
            cell.font = _tfont(bold=True)
            cell.alignment = _center() if c > 2 else _left()
            if c == 4 or (c >= 6 and c % 2 == 0):
                cell.number_format = '#,##0.00'
            if c == 3 or (c >= 5 and c % 2 == 1):
                cell.number_format = '#,##0'
        ws11.cell(row=sr, column=1).fill = _fill("F5CBA7")
        for c in range(2, len(reg_heads) + 1):
            ws11.cell(row=sr, column=c).fill = _fill("FDEBD0")

        # Column widths
        ws11.column_dimensions["A"].width = 16
        ws11.column_dimensions["B"].width = 38
        ws11.column_dimensions["C"].width = 10
        ws11.column_dimensions["D"].width = 14
        for i in range(5, len(reg_heads) + 1):
            ws11.column_dimensions[get_column_letter(i)].width = 12
        ws11.freeze_panes = "E3"
        ws11.auto_filter.ref = f"A2:{get_column_letter(len(reg_heads))}{ws11.max_row}"

    # ══════════════════════════════════════════════════════════════════════════
    # SHEET 12: СКИДКИ FEE ПО ЗАКАЗАМ (вчера/позавчера)
    # ══════════════════════════════════════════════════════════════════════════
    if fee_detail:
        ws12 = wb.create_sheet("Скидки FEE по заказам")
        ws12.sheet_properties.tabColor = "8E44AD"

        fee_heads = [
            "Дата", "№ заказа", "SKU", "Товар", "Кол-во\nшт.", "Цена,\n₽",
            "Полная\nкомиссия,₽", "Скидка\nза акцию,₽", "Скидка\n%",
            "Прочие\nкоррект.,₽", "Итог за\nвитрину,₽",
        ]
        n_fee_cols = len(fee_heads)

        ws12.merge_cells(f"A1:{get_column_letter(n_fee_cols)}1")
        _fee_dates = [r["date"] for r in fee_detail if r.get("date")]
        period_str2 = (f"{min(_fee_dates)} — {max(_fee_dates)}" if _fee_dates
                       else f"{day_before.strftime('%d.%m.%Y')} — {yesterday.strftime('%d.%m.%Y')}")
        ws12["A1"] = f"СКИДКИ ЗА УЧАСТИЕ В АКЦИЯХ (Размещение на витрине) · {period_str2}"
        ws12["A1"].font = Font(bold=True, color="FFFFFF", size=14)
        ws12["A1"].fill = _fill("8E44AD")
        ws12["A1"].alignment = Alignment(horizontal="center")
        ws12.row_dimensions[1].height = 36

        hdr_row = 2
        for c, h in enumerate(fee_heads, 1):
            ws12.cell(row=hdr_row, column=c, value=h)
        _style_header_row(ws12, hdr_row, n_fee_cols, "8E44AD")
        ws12.row_dimensions[hdr_row].height = 44

        # Сортировка по ДАТЕ по убыванию (свежие сверху); внутри даты — по артикулу/заказу.
        # Стабильная двухступенчатая сортировка: сначала sku/order ↑, затем дата ↓.
        sorted_rows = sorted(fee_detail, key=lambda x: (x["sku"], x["order_id"]))
        sorted_rows = sorted(sorted_rows, key=lambda x: x["date"], reverse=True)

        data_start = hdr_row + 1
        total_full = total_discount = total_paid = total_other = 0.0
        for row in sorted_rows:
            r = ws12.max_row + 1
            fee_full = row["fee_full"]
            fee_discount = row["fee_discount"]
            fee_paid = row["fee_paid"]
            other_adj = row.get("other_adj", 0.0)
            total_full += fee_full
            total_discount += fee_discount
            total_paid += fee_paid
            total_other += other_adj
            discount_pct = round(fee_discount / fee_full * 100, 1) if fee_full else 0.0

            ws12.cell(row=r, column=1, value=row["date"])
            order_id = row["order_id"]
            ws12.cell(row=r, column=2, value=int(order_id) if order_id else "")
            ws12.cell(row=r, column=3, value=row["sku"])
            ws12.cell(row=r, column=4, value=(row["name"] or "")[:55])
            ws12.cell(row=r, column=5, value=row["qty"])
            ws12.cell(row=r, column=6, value=row["price"])
            ws12.cell(row=r, column=7, value=fee_full)
            ws12.cell(row=r, column=8, value=fee_discount)
            ws12.cell(row=r, column=9, value=discount_pct)
            ws12.cell(row=r, column=10, value=other_adj)
            ws12.cell(row=r, column=11, value=fee_paid)

            for c in range(1, n_fee_cols + 1):
                cell = ws12.cell(row=r, column=c)
                cell.border = _border()
                cell.alignment = _center() if c not in (3, 4) else _left()
                if c in (6, 7, 8, 10, 11):
                    cell.number_format = '#,##0.00'
                if c == 5:
                    cell.number_format = '#,##0'
                if c == 9:
                    cell.number_format = '0.0"%"'

            # Heatmap: flag orders where the joint-promo discount disappeared
            status_cell = ws12.cell(row=r, column=8)
            pct_cell = ws12.cell(row=r, column=9)
            if fee_full > 0 and fee_discount == 0:
                status_cell.fill = _fill(RED_LIGHT)
                status_cell.font = _tfont(bold=True, color="B71C1C")
                pct_cell.fill = _fill(RED_LIGHT)
                pct_cell.font = _tfont(bold=True, color="B71C1C")
            elif fee_full > 0 and discount_pct < 90:
                status_cell.fill = _fill(ORANGE_LT)
                pct_cell.fill = _fill(ORANGE_LT)
            elif fee_full > 0:
                status_cell.fill = _fill(GREEN_LT)
                pct_cell.fill = _fill(GREEN_LT)

            # Flag "прочие корректировки" (штрафы за просрочку, скидка за лояльность и т.п.)
            if other_adj != 0:
                adj_cell = ws12.cell(row=r, column=10)
                if other_adj > 0:
                    adj_cell.fill = _fill(ORANGE_LT)
                    adj_cell.font = _tfont(bold=True, color="E65100")
                else:
                    adj_cell.fill = _fill(YELLOW_LT)
                    adj_cell.font = _tfont(bold=True, color="F57F17")

        last_data_row = ws12.max_row

        # Totals row
        sr = last_data_row + 1
        ws12.cell(row=sr, column=1, value="ИТОГО").font = _tfont(sz=12, bold=True)
        ws12.cell(row=sr, column=7, value=round(total_full, 2))
        ws12.cell(row=sr, column=8, value=round(total_discount, 2))
        ws12.cell(row=sr, column=9,
            value=round(total_discount / total_full * 100, 1) if total_full else 0.0)
        ws12.cell(row=sr, column=10, value=round(total_other, 2))
        ws12.cell(row=sr, column=11, value=round(total_paid, 2))
        for c in range(1, n_fee_cols + 1):
            cell = ws12.cell(row=sr, column=c)
            cell.border = _border()
            cell.font = _tfont(bold=True)
            cell.alignment = _center() if c not in (3, 4) else _left()
            cell.fill = _fill("E8DAEF")
            if c in (6, 7, 8, 10, 11):
                cell.number_format = '#,##0.00'
            if c == 9:
                cell.number_format = '0.0"%"'

        # Column widths
        ws12.column_dimensions["A"].width = 12
        ws12.column_dimensions["B"].width = 16
        ws12.column_dimensions["C"].width = 16
        ws12.column_dimensions["D"].width = 50
        for letter in ("E", "F", "G", "H", "I", "J", "K"):
            ws12.column_dimensions[letter].width = 13
        ws12.freeze_panes = "E3"
        ws12.auto_filter.ref = f"A2:{get_column_letter(n_fee_cols)}{last_data_row}"

        # Пояснение
        note_row = sr + 2
        ws12.cell(row=note_row, column=1,
            value="«Прочие коррект.» = АХ − (Полная комиссия − Скидка за акцию): штрафы за "
                  "просрочку отправки/доставки (+), скидка за лояльность или индивидуальная "
                  "скидка (−). Не относится к акционной скидке.")
        ws12.cell(row=note_row, column=1).font = _tfont(color="555555")
        ws12.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=n_fee_cols)

    # ══════════════════════════════════════════════════════════════════════════
    # НОВЫЕ СРЕЗЫ: индекс цен, дырявая воронка, качество карточек + отзывы
    # ══════════════════════════════════════════════════════════════════════════
    price_reco = kwargs.get("price_reco", {}) or {}
    reviews = kwargs.get("reviews", {}) or {}
    name_by_sku = {o["shopSku"]: (o.get("name") or o["shopSku"]) for o in all_offers}
    cr_by_sku = {o["shopSku"]: o.get("content_rating") for o in all_offers}
    cs_by_sku = {o["shopSku"]: o.get("card_status", "") for o in all_offers}

    # ── ЛИСТ: ИНДЕКС ЦЕН ───────────────────────────────────────────────────────
    if price_reco:
        wsp = wb.create_sheet("Индекс цен")
        wsp.sheet_view.showGridLines = False
        from collections import Counter as _Counter
        bands = _Counter(v["band"] for v in price_reco.values())
        n_good = bands.get("Выгодная", 0); n_mod = bands.get("Умеренная", 0)
        n_high = bands.get("Завышенная", 0)
        wsp.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
        wsp["A1"] = (f"Привлекательность цен (раздел «Цены» кабинета). "
                     f"🟢 Выгодных: {n_good} · 🟡 Умеренных: {n_mod} · 🔴 Завышенных: {n_high}. "
                     f"«Выгодная» цена ≤ порога — лучше ранжируется и попадает в подборки. "
                     f"Привязка к индексу видимости: завышенная цена → меньше показов.")
        wsp["A1"].font = _tfont(sz=10, color="555555"); wsp["A1"].alignment = _left()
        heads = ["Товар", "SKU", "Ваша\nцена ₽", "Привлекат.", "Порог\nвыгодной ₽",
                 "Порог\nсредней ₽", "Показы\n(реко)", "Рекомендация"]
        hr = 2
        for c, h in enumerate(heads, 1):
            wsp.cell(row=hr, column=c, value=h)
        _style_header_row(wsp, hr, len(heads))
        wsp.freeze_panes = "C3"
        wsp.column_dimensions["A"].width = 40
        wsp.column_dimensions["B"].width = 14
        for c in range(3, 8):
            wsp.column_dimensions[get_column_letter(c)].width = 12
        wsp.column_dimensions["H"].width = 38
        band_order = {"Завышенная": 0, "Умеренная": 1, "Выгодная": 2, "—": 3}
        rows_p = sorted(price_reco.items(),
                        key=lambda kv: (band_order.get(kv[1]["band"], 3), -(kv[1].get("shows") or 0)))
        r = hr + 1
        for sku, v in rows_p:
            band = v["band"]; price = v.get("price"); opt = v.get("optimal_price")
            avg = v.get("average_price")
            if band == "Завышенная":
                reco = f"Снизить до ≤{opt:.0f}₽ — цена выше рынка, теряете показы" if opt else "Снизить цену — выше рынка, но порог не рассчитан"
            elif band == "Умеренная":
                reco = (f"Снизить до {opt:.0f}₽ → попадёте в «выгодные», больше показов" if opt
                        else f"Умеренная — порог «выгодной» не рассчитан Яндексом (мало данных по конкурентам){f', ср.рынок {avg:.0f}₽' if avg else ''}")
            elif band == "Выгодная":
                reco = (f"OK — конкурентна. Цель {opt:.0f}₽ для запаса" if opt else "OK — цена конкурентна")
            else:
                reco = "Нет оценки (мало трафика по модели)"
            wsp.cell(row=r, column=1, value=name_by_sku.get(sku, sku)[:55])
            wsp.cell(row=r, column=2, value=sku)
            wsp.cell(row=r, column=3, value=price)
            bc = wsp.cell(row=r, column=4, value=band)
            wsp.cell(row=r, column=5, value=opt)
            wsp.cell(row=r, column=6, value=avg)
            wsp.cell(row=r, column=7, value=v.get("shows") or None)
            wsp.cell(row=r, column=8, value=reco)
            bc.fill = _fill(GREEN_LT if band == "Выгодная" else
                            YELLOW_LT if band == "Умеренная" else
                            RED_LIGHT if band == "Завышенная" else GREY_LT)
            bc.font = _tfont(bold=True)
            for c in range(1, len(heads) + 1):
                cell = wsp.cell(row=r, column=c); cell.border = _border()
                cell.alignment = _left() if c in (1, 8) else _center() if c > 2 else _left()
                if c in (3, 5, 6):
                    cell.number_format = '#,##0'
            r += 1
        wsp.auto_filter.ref = f"A{hr}:{get_column_letter(len(heads))}{r-1}"

    # ── ЛИСТ: ДЫРЯВАЯ ВОРОНКА ──────────────────────────────────────────────────
    # Агрегируем воронку за период по каждому SKU из df_full.
    if not df_full.empty and "shopSku" in df_full.columns:
        def _colsum(g, col):
            return float(g[col].fillna(0).sum()) if col in g.columns else 0.0
        funnel_rows = []
        for sku, g in df_full.groupby("shopSku"):
            shows = _colsum(g, "shows"); carts = _colsum(g, "carts")
            orders = _colsum(g, "orders"); revenue = _colsum(g, "revenue")
            if shows < 1:
                continue
            _pct2 = lambda a, b: round(a / b * 100, 2) if b else None
            s2c = _pct2(carts, shows)         # показ → корзина (мелкие доли → 2 знака)
            c2o = _pct(orders, carts)         # корзина → заказ
            s2o = _pct2(orders, shows)        # показ → заказ (мелкие доли → 2 знака)
            funnel_rows.append({
                "sku": sku, "shows": shows, "carts": carts, "orders": orders,
                "revenue": revenue, "s2c": s2c, "c2o": c2o, "s2o": s2o,
            })
        if funnel_rows:
            shows_list = sorted(x["shows"] for x in funnel_rows)
            median_shows = shows_list[len(shows_list) // 2]
            # «дырявый» = трафик выше медианы И конверсия корзина→заказ < 30%
            for x in funnel_rows:
                x["leaky"] = (x["shows"] >= max(median_shows, 500)
                              and (x["c2o"] is not None and x["c2o"] < 30))
            wsf = wb.create_sheet("Дырявая воронка")
            wsf.sheet_view.showGridLines = False
            n_leaky = sum(1 for x in funnel_rows if x["leaky"])
            wsf.merge_cells(start_row=1, start_column=1, end_row=1, end_column=8)
            wsf["A1"] = (f"Воронка за период {ue_date_from}–{ue_date_to}. 🔴 проблемных SKU: {n_leaky} "
                         f"(трафик выше медианы + конверсия корзина→заказ < 30%). "
                         f"Высокий трафик при низкой конверсии = деньги на ветер: смотри цену, "
                         f"отзывы, контент карточки.")
            wsf["A1"].font = _tfont(sz=10, color="555555"); wsf["A1"].alignment = _left()
            heads = ["Товар", "SKU", "Показы", "В корзину", "Заказы",
                     "Показ→\nкорзина %", "Корзина→\nзаказ %", "Показ→\nзаказ %"]
            hr = 2
            for c, h in enumerate(heads, 1):
                wsf.cell(row=hr, column=c, value=h)
            _style_header_row(wsf, hr, len(heads))
            wsf.freeze_panes = "C3"
            wsf.column_dimensions["A"].width = 40
            wsf.column_dimensions["B"].width = 14
            for c in range(3, 9):
                wsf.column_dimensions[get_column_letter(c)].width = 12
            # сортировка: дырявые сверху, потом по показам
            funnel_rows.sort(key=lambda x: (not x["leaky"], -x["shows"]))
            r = hr + 1
            for x in funnel_rows:
                wsf.cell(row=r, column=1, value=name_by_sku.get(x["sku"], x["sku"])[:55])
                wsf.cell(row=r, column=2, value=x["sku"])
                wsf.cell(row=r, column=3, value=int(x["shows"]))
                wsf.cell(row=r, column=4, value=int(x["carts"]))
                wsf.cell(row=r, column=5, value=int(x["orders"]))
                wsf.cell(row=r, column=6, value=x["s2c"])
                c2o_cell = wsf.cell(row=r, column=7, value=x["c2o"])
                wsf.cell(row=r, column=8, value=x["s2o"])
                if x["leaky"]:
                    for c in range(1, 3):
                        wsf.cell(row=r, column=c).fill = _fill(RED_LIGHT)
                    c2o_cell.fill = _fill(RED_LIGHT); c2o_cell.font = _tfont(bold=True, color="B71C1C")
                elif x["c2o"] is not None and x["c2o"] >= 50:
                    c2o_cell.fill = _fill(GREEN_LT)
                for c in range(1, len(heads) + 1):
                    cell = wsf.cell(row=r, column=c); cell.border = _border()
                    cell.alignment = _left() if c == 1 else _center() if c > 2 else _left()
                    if c in (3, 4, 5):
                        cell.number_format = '#,##0'
                    if c in (6, 8):           # показ→корзина, показ→заказ — 2 знака
                        cell.number_format = '0.00"%"'
                    if c == 7:                # корзина→заказ — крупнее, 1 знак
                        cell.number_format = '0.0"%"'
                r += 1
            wsf.auto_filter.ref = f"A{hr}:{get_column_letter(len(heads))}{r-1}"

    # ── ЛИСТ: КАЧЕСТВО КАРТОЧЕК + ОТЗЫВЫ ───────────────────────────────────────
    skus_q = set(s for s, cr in cr_by_sku.items() if cr is not None) | set(reviews.keys())
    if skus_q:
        wsq = wb.create_sheet("Качество и отзывы")
        wsq.sheet_view.showGridLines = False
        rated = [cr for cr in cr_by_sku.values() if cr is not None]
        avg_cr = round(sum(rated) / len(rated)) if rated else 0
        total_unans = sum(v["unanswered"] for v in reviews.values())
        # Статус карточки → человекочитаемая подпись.
        STATUS_LABELS = {
            "HAS_CARD_CAN_UPDATE":        "Есть, можно улучшить",
            "HAS_CARD_CAN_UPDATE_ERRORS": "Есть, ошибки заполнения!",
            "HAS_CARD_NO_UPDATE":         "Есть, изменения недоступны",
            "NO_CARD_ERRORS":             "Нет карточки / ошибки",
            "NO_CARD_NEED_CONTENT":       "Нет карточки — нужен контент",
        }
        def _status_label(s):
            return STATUS_LABELS.get(s, s)
        # Колонку «Статус карточки» показываем только если разных статусов > 1.
        statuses = set(cs_by_sku.get(s, "") for s in skus_q if cs_by_sku.get(s))
        show_status = len(statuses) > 1
        single_status = _status_label(next(iter(statuses))) if len(statuses) == 1 else ""
        cols = [("Товар", "name", 40), ("SKU", "sku", 14), ("Контент-\nрейтинг", "cr", 10)]
        if show_status:
            cols.append(("Статус\nкарточки", "status", 20))
        cols += [("Ср.\nрейтинг", "ar", 10), ("Отзывов", "cnt", 10),
                 ("Без\nответа", "unans", 10), ("5★/4★/3★/2★/1★", "dist", 16),
                 ("Посл.\nотзыв", "last", 10), ("Послед. негатив", "neg", 42)]
        ncol = len(cols)
        col_idx = {key: i for i, (_, key, _) in enumerate(cols, 1)}
        wsq.merge_cells(start_row=1, start_column=1, end_row=1, end_column=ncol)
        status_note = (f" Статус всех карточек: «{single_status}»." if single_status else "")
        wsq["A1"] = (f"Контент-рейтинг карточек (0–100, ср. {avg_cr}) и отзывы. "
                     f"Без ответа продавца: {total_unans} отзыв(ов). "
                     f"Низкий контент-рейтинг и плохие отзывы душат конверсию из воронки. "
                     f"Сортировка: худшие карточки сверху.{status_note}")
        wsq["A1"].font = _tfont(sz=10, color="555555"); wsq["A1"].alignment = _left()
        hr = 2
        for c, (h, _, w) in enumerate(cols, 1):
            wsq.cell(row=hr, column=c, value=h)
            wsq.column_dimensions[get_column_letter(c)].width = w
        _style_header_row(wsq, hr, ncol)
        wsq.freeze_panes = "C3"
        # сортируем: сначала по контент-рейтингу (худшие сверху), None в конец
        def _qkey(s):
            cr = cr_by_sku.get(s)
            return (cr is None, cr if cr is not None else 999)
        last_key = cols[-1][1]
        r = hr + 1
        for sku in sorted(skus_q, key=_qkey):
            cr = cr_by_sku.get(sku)
            rv = reviews.get(sku, {})
            dist = ""
            if rv.get("count"):
                dist = f"{rv.get('r5',0)}/{rv.get('r4',0)}/{rv.get('r3',0)}/{rv.get('r2',0)}/{rv.get('r1',0)}"
            raw_status = cs_by_sku.get(sku, "")
            vals = {
                "name": name_by_sku.get(sku, sku)[:55], "sku": sku, "cr": cr,
                "status": _status_label(raw_status),
                "ar": rv.get("avg_rating") or None, "cnt": rv.get("count") or None,
                "unans": rv.get("unanswered") or None, "dist": dist or None,
                "last": rv.get("last_date") or None, "neg": rv.get("last_neg") or None,
            }
            for key, ci in col_idx.items():
                wsq.cell(row=r, column=ci, value=vals[key])
            if show_status and ("ERROR" in raw_status or raw_status.startswith("NO_CARD")):
                sc = wsq.cell(row=r, column=col_idx["status"])
                sc.fill = _fill(RED_LIGHT); sc.font = _tfont(bold=True, color="B71C1C")
            if cr is not None:
                crc = wsq.cell(row=r, column=col_idx["cr"])
                crc.fill = _fill(RED_LIGHT if cr < 50 else YELLOW_LT if cr < 80 else GREEN_LT)
                crc.font = _tfont(bold=True)
            ar = rv.get("avg_rating")
            if ar:
                wsq.cell(row=r, column=col_idx["ar"]).fill = _fill(
                    RED_LIGHT if ar < 4 else YELLOW_LT if ar < 4.5 else GREEN_LT)
            if rv.get("unanswered"):
                uc = wsq.cell(row=r, column=col_idx["unans"])
                uc.fill = _fill(ORANGE_LT); uc.font = _tfont(bold=True, color="E65100")
            for c in range(1, ncol + 1):
                cell = wsq.cell(row=r, column=c); cell.border = _border()
                cell.alignment = _left() if c in (1, col_idx[last_key]) else _center() if c > 2 else _left()
            r += 1
        wsq.auto_filter.ref = f"A{hr}:{get_column_letter(ncol)}{r-1}"

    # ══════════════════════════════════════════════════════════════════════════
    # ЛИСТ: СВЕРКА / ВЫПЛАТЫ (что реально перевёл Яндекс по артикулам, FBS/FBY)
    # ══════════════════════════════════════════════════════════════════════════
    # Импорт явный: раньше имя приезжало из секции «Юнит-экономика», а она под
    # `if all_fee_orders:` — без заказов вся сверка падала на NameError.
    from .cost_prices import get_cost_price
    netting = kwargs.get("netting", {}) or {}
    netting_from = kwargs.get("netting_from")
    netting_to = kwargs.get("netting_to")
    # Квартальный нетинг (90д) — устойчивая маржа; агрегируем по SKU (без FBS/FBY).
    netting_quarter = kwargs.get("netting_quarter", {}) or {}
    nq_from = kwargs.get("netting_quarter_from")
    nq_to = kwargs.get("netting_quarter_to")
    q_by_sku = {}
    for (qsku, _qmodel), qa in netting_quarter.items():
        acc = q_by_sku.setdefault(qsku, {"qty": 0.0, "ret_qty": 0.0, "buyer_pay": 0.0,
                                         "payout": 0.0, "returns": 0.0,
                                         "promo": 0.0, "points": 0.0})
        for k in ("qty", "ret_qty", "buyer_pay", "payout", "returns", "promo", "points"):
            acc[k] += qa.get(k, 0)

    # Последние N дней (по умолчанию 7) — «в моменте», без сглаживания. Не путать с
    # Σ 90д: тут сознательно НЕ усредняем, если данных мало — цифра просто шумная/0.
    netting_recent = kwargs.get("netting_recent", {}) or {}
    netting_recent_days = kwargs.get("netting_recent_days", 7)
    recent_by_sku = {}
    for (rsku, _rmodel), ra in netting_recent.items():
        acc = recent_by_sku.setdefault(rsku, {"qty": 0.0, "ret_qty": 0.0, "buyer_pay": 0.0,
                                              "payout": 0.0, "returns": 0.0})
        for k in ("qty", "ret_qty", "buyer_pay", "payout", "returns"):
            acc[k] += ra.get(k, 0)

    def _profit_pu(agg_for_sku, sku_for):
        """Чистая прибыль/шт из компактного агрегата {qty,ret_qty,buyer_pay,payout,returns}.
        0, если данных за период нет (напр. 0 продаж за последние 7 дней)."""
        if not agg_for_sku or agg_for_sku.get("qty", 0) <= 0:
            return 0.0
        qty = agg_for_sku["qty"]; net_q = qty - agg_for_sku.get("ret_qty", 0)
        if net_q <= 0:
            return 0.0
        cost = (get_cost_price(sku_for) or 0) * net_q
        tax = (agg_for_sku["buyer_pay"] + agg_for_sku.get("returns", 0)) * 0.07
        profit = agg_for_sku["payout"] - cost - tax
        return round(profit / net_q, 2)

    # ── ВЗВЕШЕННАЯ ₽/шт: бакеты 0–30д и 31–90д, вес по объёму И свежести ────────
    # Окна нетинга ВЛОЖЕНЫ (7д ⊂ 30д ⊂ 90д) — усреднять их напрямую нельзя (тройной
    # счёт свежих дней). Поэтому режем на НЕПЕРЕСЕКАЮЩИЕСЯ бакеты: 0–30д и 31–90д
    # (= 90д − 30д покомпонентно; законно, т.к. формула прибыли линейна по агрегату).
    # 7-дневное окно в расчёт НЕ входит осознанно: замер по каталогу показал, что
    # оно смещено вниз на ~122 ₽/шт (медиана) при разбросе σ≈263 ₽/шт и 68% SKU ниже
    # квартальной базы. Это артефакт ЛАГА ВЫПЛАТ (штуки считаются по «Платёж
    # покупателя», а услуги за те же заказы удерживаются неделями позже, т.е. на
    # коротком окне payments и services относятся к РАЗНЫМ заказам), а не экономика.
    # Класть такую подвыборку в оценку с любым весом — значит систематически занижать.
    W_RECENT, W_OLD = 2.0, 1.0   # 2:1 ≈ экспоненц. затухание с полупериодом ~45 дней
    THIN_UNITS = 10              # меньше — цифра ненадёжна, красим серым
    NETT_FIELDS = ("qty", "ret_qty", "buyer_pay", "payout", "returns")

    n30_by_sku = {}
    for (s30, _m30), a30 in netting.items():
        acc = n30_by_sku.setdefault(s30, {f: 0.0 for f in NETT_FIELDS})
        for f in NETT_FIELDS:
            acc[f] += a30.get(f, 0)

    def _profit_of(agg_for_sku, sku_for):
        """(прибыль ₽, чистые штуки) из сырого агрегата. Та же формула, что у Σ 90д;
        линейна по агрегату → применима и к РАЗНОСТИ периодов (бакет 31–90д)."""
        if not agg_for_sku:
            return 0.0, 0.0
        net_q = agg_for_sku.get("qty", 0) - agg_for_sku.get("ret_qty", 0)
        if net_q <= 0:
            return 0.0, 0.0
        cost = (get_cost_price(sku_for) or 0) * net_q
        tax = (agg_for_sku.get("buyer_pay", 0) + agg_for_sku.get("returns", 0)) * 0.07
        return agg_for_sku.get("payout", 0) - cost - tax, net_q

    def _old_bucket(a_q, a_new):
        """Бакет 31–90д = 90д − 30д покомпонентно. None, если разность невалидна:
        90-дневный кэш живёт 7 дней, и его окно заканчивается РАНЬШЕ 30-дневного
        (у 30д кэш на сутки) → у части SKU штуки в разности уходят в ноль/минус.
        Такой бакет отбрасываем целиком — оценка строится на одном свежем окне.
        Деньги НЕ клампим: returns/payout законно отрицательные, обнуление их
        сломало бы прибыль."""
        if not a_q:
            return None
        diff = {f: a_q.get(f, 0) - (a_new.get(f, 0) if a_new else 0) for f in NETT_FIELDS}
        if diff["qty"] < 0 or (diff["qty"] - diff["ret_qty"]) <= 0:
            return None
        return diff

    blend_by_sku = {}   # sku -> (взвеш. ₽/шт, тренд ₽/шт|None, чист. штук в расчёте)
    blend_tot_num = blend_tot_den = 0.0   # для строки ИТОГО (агрегат, не среднее средних)
    for _s in set(n30_by_sku) | set(q_by_sku):
        a_new = n30_by_sku.get(_s)
        a_old = _old_bucket(q_by_sku.get(_s), a_new)
        p_new, n_new = _profit_of(a_new, _s)
        p_old, n_old = _profit_of(a_old, _s)
        W = W_RECENT * n_new + W_OLD * n_old
        if W <= 0:
            continue
        blended = (W_RECENT * p_new + W_OLD * p_old) / W
        blend_tot_num += W_RECENT * p_new + W_OLD * p_old
        blend_tot_den += W
        # Тренд в ₽/шт, а НЕ в %: при базе ≈0 проценты взрываются на порядки
        # (реальный пример — +3260% там, где сдвиг всего +157 ₽/шт).
        trend = ((p_new / n_new) - (p_old / n_old)) if (n_new and n_old) else None
        blend_by_sku[_s] = (round(blended, 2),
                            round(trend, 1) if trend is not None else None,
                            n_new + n_old)

    if netting:
        wsn = wb.create_sheet("Сверка выплат")
        ozon_ppu, ozon_date = load_ozon_ppu(report_folder)
        wsn.sheet_properties.tabColor = "C0392B"
        period = (f"{netting_from.strftime('%d.%m')}–{netting_to.strftime('%d.%m.%Y')}"
                  if netting_from and netting_to else "")
        nnote = (f"РЕАЛЬНЫЕ выплаты от Яндекса по артикулам (отчёт «По платежам») за {period}. "
                 f"К перечислению ФАКТ = Платежи покуп. − Услуги − Возвраты − Скидка акций + Баллы + Прочее "
                 f"(все статьи — отдельные проводки Яндекса, в сумме = реальные деньги на счёт). "
                 f"«Удержано услуги» = «Оплата услуг Яндекс.Маркета» (комиссия+буст+логистика+хранение). "
                 f"«Скидка акций» (−) — твоя скидка по акции (реально уменьшает). «Баллы-компенсация» (+) — "
                 f"Яндекс возвращает часть скидки баллами. Их разница = твоя непокрытая доля скидки. "
                 f"«Прочее» — контроль, должно быть ≈0. «Чистая прибыль ФАКТ» = к перечислению − себест(на ЧИСТЫЕ "
                 f"штуки, продано−возвраты) − налог 7% (с нетто-выручки). Возвраты возвращаются годными на склад, "
                 f"поэтому их себест/налог НЕ списываем; реальная цена возврата — только логистика туда-обратно (в «Услугах»). "
                 f"«ФАКТ ₽/шт» = прибыль / (Продано − Возвраты), т.е. на чистую (выкупленную) штуку. "
                 f"Выплаты смещены: оплаты приходят через дни/недели после заказа — поэтому дневной "
                 f"₽/шт скачет. Оранжевая строка «Σ 90д» под каждым SKU — устойчивая КВАРТАЛЬНАЯ "
                 f"чистая прибыль (без разбивки FBS/FBY), где смещение усредняется; на неё и ориентируйся. "
                 f"«ФАКТ ₽/шт (посл. {netting_recent_days}д)» — СИГНАЛ в моменте, а НЕ оценка прибыльности: "
                 f"на таком коротком окне выплаты и удержания относятся к разным заказам, поэтому цифра "
                 f"смещена вниз (медиана −122 ₽/шт к кварталу, разброс σ≈263) — по ней смотрят «что-то "
                 f"изменилось», но не «сколько зарабатываю»; 0 = продаж за период не было. "
                 f"⭐ «ВЗВЕШ. ₽/шт» — ГЛАВНАЯ цифра «сколько артикул реально приносит сейчас»: два "
                 f"НЕПЕРЕСЕКАЮЩИХСЯ бакета (0–30д и 31–90д) с весом 2:1 по свежести и с весом по объёму "
                 f"продаж (≈ затухание с полупериодом 45 дней). Окна нетинга вложены (7д⊂30д⊂90д), поэтому "
                 f"усреднять их напрямую нельзя — тройной счёт; 7д в расчёт не входит осознанно (см. выше). "
                 f"СЕРАЯ цифра = меньше {THIN_UNITS} чистых штук за квартал, статистике не верить. "
                 f"«Тренд ₽/шт» = ₽/шт(0–30д) − ₽/шт(31–90д): + значит маржа артикула растёт, − падает "
                 f"(в рублях, а не в %, т.к. при базе около нуля проценты взлетают на порядки). "
                 f"ДВЕ ПОСЛЕДНИЕ КОЛОНКИ — это ЗАМКНУТЫЙ КОНТУР БОНУСНОГО БАЛАНСА, а не движение "
                 f"живых денег: у обеих в отчёте статус начинается со слова «Справочно» "
                 f"(«Справочно: пополнен баланс» / «Справочно: списан с баланса»), тогда как у реальных "
                 f"денег статусы другие («Переведён по графику выплат», «Удержан из платежей покупателей»). "
                 f"Механика: Яндекс САМ финансирует скидку покупателю → пополняет твой бонусный баланс "
                 f"баллами → тут же списывает с этого баланса оплату твоих услуг. "
                 f"🟢 «БАЛЛЫ ₽/шт» (+): «Баллы за скидку Маркета/Плюс» — сколько Яндекс положил на баланс. "
                 f"Привязаны К ТОВАРУ (кол-во в проводках = штукам). "
                 f"🔴 «УСЛУГИ ЗА СЧЁТ БАЛЛОВ ₽/шт» (−): строка «Скидка за участие в совместных акциях» — "
                 f"название обманчивое, по факту это ТВОИ УСЛУГИ, оплаченные с бонусного баланса "
                 f"(размещение ~81%, буст ~11%, логистика ~8%). Кол-во во ВСЕХ таких проводках = 0, "
                 f"к цене товара отношения НЕ имеет — поэтому строка бывает и у артикула ВНЕ всяких акций "
                 f"(у него тоже есть услуги); на штуки делим только для сопоставимости. "
                 f"РАЗНИЦА ОСМЫСЛЕННА НА УРОВНЕ КАТАЛОГА, А НЕ ОТДЕЛЬНОГО SKU: по каталогу Яндекс кладёт "
                 f"на баланс на 6–10% больше, чем списывает — "
                 f"т.е. участие в акциях в целом бесплатно и слегка в плюс. Но по одному артикулу разница "
                 f"НЕНАДЁЖНА: баллы считаются по товару, а списание — по услугам, с собственным расчётным "
                 f"периодом, поэтому половинки пары часто попадают в разные окна (отсюда разброс от −83 "
                 f"до +94 ₽/шт — это рассинхрон учёта, а не разная экономика товаров). "
                 f"ВАЖНО: само снижение цены по акции отдельной строкой в отчёте НЕ проходит — "
                 f"покупатель просто платит меньше, и это уже сидит в «Платежах покуп.», "
                 f"а значит и в «Чистой прибыли ФАКТ». Отдельной цифры «во что обошлась скидка» "
                 f"не существует: это недополученная выручка, а не расход.")
        nnote += (f" СРАВНЕНИЕ С ОЗОНОМ (колонки рядом с ⭐ВЗВЕШ): «ОЗОН ₽/ед» — прибыль на штуку "
                  f"из отчёта ozon_report_{ozon_date}, «Δ ЯМ−Озон» — насколько ⭐ВЗВЕШ Яндекса "
                  f"выше (+) или ниже (−) Озона. Артикулы сопоставляются точным совпадением; "
                  f"прочерк — на Озоне этого артикула нет. Методики площадок разные (у Озона "
                  f"данные из выгрузки начислений, у Яндекса — взвешенная по 90 дням), поэтому "
                  f"смотреть на знак и порядок разницы, а не на точные рубли."
                  if ozon_ppu else " Сравнение с Озоном недоступно: отчёт ozon_report_*.xlsx не найден.")
        wsn["A1"] = nnote
        wsn["A1"].font = _tfont(sz=10, color="555555"); wsn["A1"].alignment = _left()
        wsn.row_dimensions[1].height = 100

        nheads = ["Товар", "SKU", "Модель", "Продано\nшт.", "Возвр.\nшт.",
                  "Платежи\nпокуп. ₽", "Удержано\nуслуги ₽", "Возвраты ₽",
                  "Скидка\nакций ₽", "Баллы-\nкомпенс. ₽", "Прочее ₽",
                  "К перечисл.\nФАКТ ₽", "Себест. ₽", "Налог\n7% ₽",
                  "Чистая\nприбыль\nФАКТ ₽", "Чистая\nприбыль\nФАКТ ₽/шт", "Маржа %",
                  f"ФАКТ ₽/шт\n(посл. {netting_recent_days}д)",
                  "⭐ ВЗВЕШ.\n₽/шт", "Тренд\n₽/шт",
                  "🟢 БАЛЛЫ\n₽/шт", "🔴 УСЛУГИ ЗА\nСЧЁТ БАЛЛОВ\n₽/шт",
                  "ОЗОН\n₽/ед", "Δ ЯМ−Озон\n₽/шт"]
        wsn.merge_cells(f"A1:{get_column_letter(len(nheads))}1")
        # Строки собираются в ЛОГИЧЕСКОМ порядке nheads, а в лист пишутся в ФИЗИЧЕСКОМ:
        # две главные цифры — «ФАКТ ₽/шт» (16) и «⭐ ВЗВЕШ.» (19) — стоят сразу после
        # «Модели» и закреплены вместе с ней, чтобы считывались без прокрутки.
        # Одна перестановка на все типы строк (по моделям / Σ 90д / ИТОГО), а не
        # правка каждого списка значений — так колонки не могут разъехаться.
        PHYS_ORDER = [1, 2, 3, 16, 19, 23, 24, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13,
                      14, 15, 17, 18, 20, 21, 22]
        phys = {logical: pos for pos, logical in enumerate(PHYS_ORDER, 1)}

        def _reorder(lst):
            return [lst[i - 1] for i in PHYS_ORDER]

        # денежные колонки; прибыль/шт (16), (18)–(22) с копейками — номера логические
        MONEY_N = tuple(phys[c] for c in (6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 18, 19, 20,
                                          21, 22, 23, 24))
        COL_PU = phys[16]; COL_MAR = phys[17]; COL_PROF = phys[15]; COL_COSTN = phys[13]
        COL_PU_RECENT = phys[18]; COL_BLEND = phys[19]; COL_TREND = phys[20]
        COL_POINTS_PU = phys[21]; COL_PREMIUM_PU = phys[22]
        COL_OZON = phys[23]; COL_OZON_DIFF = phys[24]
        DEC_COLS = (COL_PU, COL_PU_RECENT, COL_BLEND, COL_TREND,
                    COL_POINTS_PU, COL_PREMIUM_PU, COL_OZON, COL_OZON_DIFF)

        def _ozon_vals(sku_for, ym_ppu):
            """[Озон ₽/ед, Δ ЯМ−Озон]. Сопоставление ТОЛЬКО по точному совпадению
            артикула: префиксное «похоже» уже подводило (SKUbl ≠ SKUblk), а
            здесь цена ошибки — сравнение экономики разных товаров."""
            oz = ozon_ppu.get(str(sku_for))
            if oz is None:
                return ["—", "—"]
            if not isinstance(ym_ppu, (int, float)):
                return [oz, "—"]
            return [oz, round(ym_ppu - oz, 2)]

        def _style_ozon(rr, bold=False):
            oc_ = wsn.cell(row=rr, column=COL_OZON)
            dc_ = wsn.cell(row=rr, column=COL_OZON_DIFF)
            if not isinstance(oc_.value, (int, float)):
                oc_.font = _tfont(bold=bold, color="999999")
                dc_.font = _tfont(bold=bold, color="999999")
                return
            oc_.font = _tfont(bold=bold, color="1B5E20" if oc_.value > 0 else "B71C1C")
            if isinstance(dc_.value, (int, float)):
                dc_.font = _tfont(bold=True, color="1B5E20" if dc_.value > 0 else "B71C1C")
                dc_.fill = _fill(GREEN_LT if dc_.value > 0 else RED_LIGHT)

        def _tier_fill(cell, val, bold=True, font_color=None):
            """Светофор для ₽/шт: > 100 — зелёный, 50–100 — жёлтый, < 50 — красный."""
            if not isinstance(val, (int, float)):
                return
            if val > 100:
                fill, color = GREEN_LT, "1B5E20"
            elif val >= 50:
                fill, color = YELLOW_LT, "7D6608"
            else:
                fill, color = RED_LIGHT, "B71C1C"
            cell.fill = _fill(fill)
            cell.font = _tfont(bold=bold, color=font_color or color)

        def _cofin_pu(agg):
            """(баллы ₽/шт, услуги-за-счёт-баллов ₽/шт) на ЧИСТУЮ штуку — РАЗДЕЛЬНО.

            Это ЗАМКНУТЫЙ КОНТУР БОНУСНОГО БАЛАНСА, не движение живых денег: у обеих
            статей статус в отчёте начинается со «Справочно» («пополнен баланс» /
            «списан с баланса»), тогда как у реальных денег статусы «Переведён по
            графику выплат» / «Удержан из платежей покупателей». Яндекс сам финансирует
            скидку покупателю → пополняет баланс баллами → списывает с него твои услуги.
            Проверено на сыром отчёте «По платежам» (935 строк за 30д):
              • points «Баллы за скидку Маркета/Плюс» (+) — сколько положено на баланс.
                Привязаны К ТОВАРУ, кол-во = штукам.
              • promo «Скидка за участие в совместных акциях» (−) — название обманчивое:
                это ТВОИ УСЛУГИ, оплаченные с баланса (размещение −146.9к, буст −20.2к,
                логистика −14.5к). Кол-во во ВСЕХ строках = 0, к цене товара не
                относится → бывает и у SKU вне всяких акций. На штуки делим лишь
                для сопоставимости, это не «цена за штуку».
            ⚠️ РАЗНИЦА осмысленна ПО КАТАЛОГУ (+6–10%: Яндекс кладёт больше, чем
            списывает), но НЕ по одному SKU: половинки пары считаются на разной базе и
            в разных расчётных периодах, поэтому по артикулу это в основном рассинхрон
            учёта. Ни то, ни другое НЕ есть «стоимость снижения цены»: снижение цены
            отдельной проводки не имеет, покупатель просто платит меньше (уже в buyer_pay).
            Возвращает (None, None), если ни баллов, ни списаний не было."""
            if not agg:
                return None, None
            promo = agg.get("promo", 0); points = agg.get("points", 0)
            if not promo and not points:
                return None, None
            net_q = agg.get("qty", 0) - agg.get("ret_qty", 0)
            if net_q <= 0:
                return None, None
            return round(points / net_q, 2), round(promo / net_q, 2)
        hr = 2
        for c, h in enumerate(_reorder(nheads), 1):
            wsn.cell(row=hr, column=c, value=h)
        _style_header_row(wsn, hr, len(nheads), "C0392B")
        for col in (COL_PU, COL_BLEND):   # главные колонки — своим цветом шапки
            wsn.cell(row=hr, column=col).fill = _fill("1F3A5F")
        wsn.row_dimensions[hr].height = 40

        tot = {"qty": 0.0, "retq": 0.0, "buyer": 0.0, "serv": 0.0, "ret": 0.0, "promo": 0.0,
               "points": 0.0, "other": 0.0, "payout": 0.0, "cost": 0.0, "tax": 0.0,
               "profit": 0.0}
        r = hr
        band = False          # чередование оттенка по артикулу
        prev_sku = None

        def _style_blend(rr, sku_for, bold=False):
            """Красит «ВЗВЕШ. ₽/шт» и «Тренд ₽/шт». Серый = мало штук за квартал
            (цифру НЕ подгоняем сжатием к среднему — иначе она перестанет биться
            с остальным листом; просто помечаем как ненадёжную)."""
            bl = blend_by_sku.get(sku_for)
            if not bl:
                return
            blended, trend, units = bl
            thin = units < THIN_UNITS
            bc = wsn.cell(row=rr, column=COL_BLEND)
            # светофор по порогам; у «тонкого» SKU цвет фона остаётся, а цифра серая —
            # видно и уровень, и то, что статистике по нему верить рано
            _tier_fill(bc, blended, bold=True, font_color="999999" if thin else None)
            # Тренд у «тонкого» SKU так же ненадёжен, как и сама цифра — гасим и его,
            # иначе скачок на 3 штуках читается как сигнал.
            tc = wsn.cell(row=rr, column=COL_TREND)
            if thin or trend is None:
                tc.font = _tfont(bold=bold, color="999999")
            else:
                tc.font = _tfont(bold=bold, color="1B5E20" if trend > 0 else
                                 "B71C1C" if trend < 0 else "555555")

        def _blend_vals(sku_for):
            """Значения двух новых колонок — одинаковые на всех строках SKU (как у
            колонки 7д), чтобы работали сортировка и фильтр по любой строке."""
            bl = blend_by_sku.get(sku_for)
            if not bl:
                return ["", ""]
            return [bl[0], bl[1] if bl[1] is not None else "—"]

        def _style_cofin(rr, pts_pu, prem_pu, bold=False):
            """Баллы — всегда доход (зелёный), премия — всегда расход (красный).
            Цвет тут постоянный по смыслу колонки, а не по знаку: так сразу читается,
            что слева приход, справа удержание. «—» = проводок не было."""
            pc = wsn.cell(row=rr, column=COL_POINTS_PU)
            if pts_pu is None:
                pc.font = _tfont(bold=bold, color="999999")
            else:
                pc.fill = _fill(GREEN_LT); pc.font = _tfont(bold=True, color="1B5E20")
            mc = wsn.cell(row=rr, column=COL_PREMIUM_PU)
            if prem_pu is None:
                mc.font = _tfont(bold=bold, color="999999")
            else:
                mc.fill = _fill(RED_LIGHT); mc.font = _tfont(bold=True, color="B71C1C")

        def emit_quarter_row(rr, sku_for):
            """Итоговая строка «Σ 90д» по SKU (без разбивки FBS/FBY) — устойчивая
            квартальная чистая прибыль (не скачет от смещения выплат)."""
            qa = q_by_sku.get(sku_for)
            if not qa or qa["qty"] <= 0:
                return rr
            qty = qa["qty"]; net_q = qty - qa["ret_qty"]
            buyer = qa["buyer_pay"]; payout = qa["payout"]
            # Возвраты годные → себест на ЧИСТЫЕ штуки, налог с нетто-выручки.
            cost = round((get_cost_price(sku_for) or 0) * net_q, 2)
            tax = round((buyer + qa.get("returns", 0)) * 0.07, 2)
            profit = round(payout - cost - tax, 2)
            ppu = round(profit / net_q, 2) if net_q > 0 else 0
            margin = round(profit / buyer * 100, 1) if buyer else 0
            rr += 1
            recent_pu = _profit_pu(recent_by_sku.get(sku_for), sku_for)
            # Скидку/баллы в квартальной строке показываем СУММАМИ (столбцы 9/10),
            # чтобы нетто-цифра справа была проверяема глазами прямо в строке.
            q_pts, q_prem = _cofin_pu(qa)
            vals = ["", sku_for, "Σ 90д", int(qty), int(qa["ret_qty"]), round(buyer),
                    "", "", round(qa.get("promo", 0)), round(qa.get("points", 0)), "",
                    round(payout), cost, tax, profit, ppu, margin,
                    recent_pu] + _blend_vals(sku_for) + [
                    q_pts if q_pts is not None else "—",
                    q_prem if q_prem is not None else "—"] + \
                    _ozon_vals(sku_for, _blend_vals(sku_for)[0])
            for c, v in enumerate(_reorder(vals), 1):
                cell = wsn.cell(row=rr, column=c, value=v)
                cell.border = _border(); cell.fill = _fill("FDEBD0")   # квартальная — оранж.
                cell.font = _tfont(bold=True)
                cell.alignment = _left() if c == 1 else _center() if c > 2 else _left()
                if c in MONEY_N and isinstance(v, (int, float)):
                    cell.number_format = '#,##0' if c not in DEC_COLS else '#,##0.00'
            _style_cofin(rr, q_pts, q_prem, bold=True)
            wsn.cell(row=rr, column=COL_PROF).font = _tfont(
                bold=True, color="1B5E20" if profit > 0 else "B71C1C")
            if net_q > 0:
                _tier_fill(wsn.cell(row=rr, column=COL_PU), ppu)
            rcc = wsn.cell(row=rr, column=COL_PU_RECENT)
            rcc.font = _tfont(bold=True, color="1B5E20" if recent_pu > 0 else
                              "999999" if recent_pu == 0 else "B71C1C")
            _style_blend(rr, sku_for, bold=True)
            _style_ozon(rr, bold=True)
            return rr

        # сортировка по артикулу ↑, затем по модели
        for (sku, model), a in sorted(netting.items(), key=lambda kv: (str(kv[0][0]), kv[0][1])):
            if sku != prev_sku:
                if prev_sku is not None:
                    r = emit_quarter_row(r, prev_sku)   # закрыть предыдущий SKU
                band = not band
                prev_sku = sku
            shade = GREY_LT if band else None
            qty = a.get("qty", 0)
            ret_q = a.get("ret_qty", 0)
            net_q = qty - ret_q                   # чистые штуки = продано − возвраты
            buyer = a.get("buyer_pay", 0)
            returns_rub = a.get("returns", 0)     # возврат платежа покупателю (отриц.)
            payout = a.get("payout", 0)
            cost_u = get_cost_price(sku) or 0
            # Возвраты возвращаются годными на склад → себест НЕ списываем на них
            # (только на чистые штуки), налог 7% с нетто-выручки (за вычетом возвратов).
            cost = round(cost_u * net_q, 2)
            tax = round((buyer + returns_rub) * 0.07, 2)
            profit = round(payout - cost - tax, 2)
            profit_pu = round(profit / net_q, 2) if net_q > 0 else 0   # на чистую штуку
            margin = round(profit / buyer * 100, 1) if buyer else 0
            r += 1
            recent_pu = _profit_pu(netting_recent.get((sku, model)), sku)
            vals = [name_by_sku.get(sku, sku)[:50], sku, model, int(qty), int(ret_q),
                    buyer, a.get("services", 0), a.get("returns", 0),
                    a.get("promo", 0), a.get("points", 0), a.get("other", 0), payout,
                    cost, tax, profit, profit_pu, margin,
                    recent_pu] + _blend_vals(sku)
            pts_pu, prem_pu = _cofin_pu(a)
            vals = vals + [pts_pu if pts_pu is not None else "—",
                           prem_pu if prem_pu is not None else "—"] + \
                   _ozon_vals(sku, _blend_vals(sku)[0])
            for c, v in enumerate(_reorder(vals), 1):
                cell = wsn.cell(row=r, column=c, value=v)
                cell.border = _border()
                cell.alignment = _left() if c == 1 else _center() if c > 2 else _left()
                if shade:                         # чередование строк по артикулу
                    cell.fill = _fill(shade)
                if c in MONEY_N and isinstance(v, (int, float)):
                    cell.number_format = '#,##0' if c not in DEC_COLS else '#,##0.00'
            # смысловые цвета — через ШРИФТ (чтобы не перекрывать чередование фоном);
            # исключение — две главные колонки ₽/шт: у них светофор фоном
            pc = wsn.cell(row=r, column=COL_PROF)
            pc.font = _tfont(bold=True, color="1B5E20" if profit > 0 else "B71C1C")
            if net_q > 0:                          # без чистых штук ₽/шт = 0 условный
                _tier_fill(wsn.cell(row=r, column=COL_PU), profit_pu)
            mc = wsn.cell(row=r, column=COL_MAR)
            mc.font = _tfont(bold=True, color="1B5E20" if margin >= 20 else
                             "B71C1C" if margin < 5 else "9A7D0A")
            rc = wsn.cell(row=r, column=COL_PU_RECENT)
            rc.font = _tfont(bold=True, color="1B5E20" if recent_pu > 0 else
                             "999999" if recent_pu == 0 else "B71C1C")
            if cost_u == 0:
                cc = wsn.cell(row=r, column=COL_COSTN)
                cc.font = _tfont(bold=True, color="E65100")
            _style_blend(r, sku)
            _style_cofin(r, pts_pu, prem_pu)
            _style_ozon(r)
            tot["qty"] += qty; tot["retq"] += ret_q; tot["buyer"] += buyer
            tot["serv"] += a.get("services", 0)
            tot["ret"] += a.get("returns", 0); tot["promo"] += a.get("promo", 0)
            tot["points"] += a.get("points", 0); tot["other"] += a.get("other", 0)
            tot["payout"] += payout
            tot["cost"] += cost; tot["tax"] += tax; tot["profit"] += profit

        # закрыть последний SKU квартальной строкой
        if prev_sku is not None:
            r = emit_quarter_row(r, prev_sku)

        # ИТОГО
        r += 1
        net_tot = tot["qty"] - tot["retq"]
        tot_pu = round(tot["profit"] / net_tot, 2) if net_tot > 0 else 0
        ws_tot = ["ИТОГО", "", "", int(tot["qty"]), int(tot["retq"]),
                  round(tot["buyer"]), round(tot["serv"]),
                  round(tot["ret"]), round(tot["promo"]), round(tot["points"]), round(tot["other"]),
                  round(tot["payout"]), round(tot["cost"]), round(tot["tax"]),
                  round(tot["profit"]), tot_pu,
                  round(tot["profit"]/tot["buyer"]*100, 1) if tot["buyer"] else 0, "—",
                  round(blend_tot_num / blend_tot_den, 2) if blend_tot_den else "—", "—",
                  # ИТОГО: средние по каталогу. Обманчивы (разброс по SKU огромный) —
                  # предупреждение про это стоит в шапке листа.
                  (round(tot["points"] / net_tot, 2)
                   if net_tot > 0 and tot["points"] else "—"),
                  (round(tot["promo"] / net_tot, 2)
                   if net_tot > 0 and tot["promo"] else "—"),
                  # в ИТОГО межплощадочное сравнение не считаем: средняя по разным
                  # наборам артикулов на двух площадках сравнивала бы несопоставимое
                  "—", "—"]
        for c, v in enumerate(_reorder(ws_tot), 1):
            cell = wsn.cell(row=r, column=c, value=v)
            cell.border = _border(); cell.fill = _fill(GREY_LT); cell.font = _tfont(bold=True)
            cell.alignment = _center() if c > 2 else _left()
            if c in MONEY_N and isinstance(v, (int, float)):
                cell.number_format = '#,##0' if c not in DEC_COLS else '#,##0.00'
        _tier_fill(wsn.cell(row=r, column=COL_PU), ws_tot[16 - 1])
        _tier_fill(wsn.cell(row=r, column=COL_BLEND), ws_tot[19 - 1])

        wsn.column_dimensions["A"].width = 40
        wsn.column_dimensions["B"].width = 13
        wsn.column_dimensions["C"].width = 8
        for i in range(4, len(nheads) + 1):
            wsn.column_dimensions[get_column_letter(i)].width = 12
        wsn.column_dimensions[get_column_letter(COL_POINTS_PU)].width = 12
        wsn.column_dimensions[get_column_letter(COL_PREMIUM_PU)].width = 14
        wsn.column_dimensions[get_column_letter(COL_PU)].width = 13
        wsn.column_dimensions[get_column_letter(COL_BLEND)].width = 13
        wsn.column_dimensions[get_column_letter(COL_TREND)].width = 12
        wsn.column_dimensions[get_column_letter(COL_OZON)].width = 11
        wsn.column_dimensions[get_column_letter(COL_OZON_DIFF)].width = 12
        # закреплены A–G: две главные цифры и сравнение с Озоном видны без прокрутки
        wsn.freeze_panes = f"{get_column_letter(COL_OZON_DIFF + 1)}3"
        wsn.auto_filter.ref = f"A{hr}:{get_column_letter(len(nheads))}{r-1}"

    # ══════════════════════════════════════════════════════════════════════════
    # ЛИСТ: КОРЗИНА vs ПОКУПКА (воронка от КЛИКА, а не от показа)
    # Отличие от «Дырявой воронки»: та считает от ПОКАЗОВ и потому мешает в одну
    # цифру две разные проблемы (плохой CTR и плохую карточку). Здесь база — КЛИК,
    # т.е. человек уже зашёл в карточку: видно отдельно, убеждает ли она положить
    # в корзину и доводит ли корзину до оплаты.
    # ══════════════════════════════════════════════════════════════════════════
    if not pv_clicks.empty and not pv_carts.empty and active_skus:
        rows_f = []
        for sku in active_skus:
            cl = float(pv_clicks.loc[sku].sum()) if sku in pv_clicks.index else 0.0
            ca = float(pv_carts.loc[sku].sum()) if sku in pv_carts.index else 0.0
            od = float(pv_orders.loc[sku].sum()) if sku in pv_orders.index else 0.0
            if cl < 1:
                continue
            rows_f.append({"sku": sku, "clicks": cl, "carts": ca, "orders": od,
                           "k2c": ca / cl * 100 if cl else 0,
                           "k2o": od / cl * 100 if cl else 0,
                           "c2o": od / ca * 100 if ca else None})
        # Нормы — МЕДИАНЫ по каталогу (а не выдуманные пороги): что «низко» зависит
        # от ниши, и фиксированные 30/50% в другой категории врали бы.
        med_k2c = statistics.median([x["k2c"] for x in rows_f]) if rows_f else 0
        c2o_vals = [x["c2o"] for x in rows_f if x["c2o"] is not None and x["carts"] >= 3]
        med_c2o = statistics.median(c2o_vals) if c2o_vals else 0

        for x in rows_f:
            # Потенциал: сколько заказов было бы, дотяни корзина→заказ до медианы.
            gap_ord = max(0.0, x["carts"] * med_c2o / 100 - x["orders"]) if x["carts"] else 0.0
            pu = (blend_by_sku.get(x["sku"]) or (None,))[0]
            x["lost_orders"] = gap_ord
            x["lost_rub"] = round(gap_ord * pu) if (pu and pu > 0) else None
            x["no_order_carts"] = max(0.0, x["carts"] - x["orders"])
            # Диагноз: разделяем «не кладут» и «кладут, но не платят»
            if x["clicks"] < 15:
                x["diag"] = "мало кликов — не судить"
            elif x["c2o"] is not None and x["c2o"] < med_c2o * 0.6 and x["carts"] >= 3:
                x["diag"] = "🔴 кладут, но не покупают"
            elif x["k2c"] < med_k2c * 0.6:
                x["diag"] = "🟡 карточка не убеждает"
            elif x["c2o"] is not None and x["c2o"] >= med_c2o:
                x["diag"] = "🟢 норма"
            else:
                x["diag"] = "—"
        rows_f.sort(key=lambda x: -(x["lost_rub"] or 0))

        wsk = wb.create_sheet("Корзина vs Покупка")
        wsk.sheet_properties.tabColor = "6C3483"
        knote = (f"Воронка от КЛИКА за {len(all_dates)}д (человек уже открыл карточку). "
                 f"«Клик→корзина» — убеждает ли карточка положить товар; «Корзина→заказ» — "
                 f"доводится ли корзина до оплаты. Медианы каталога: клик→корзина "
                 f"{med_k2c:.1f}%, корзина→заказ {med_c2o:.1f}% — это и есть норма, "
                 f"пороги не выдуманные. 🔴 «кладут, но не покупают» = корзина→заказ ниже "
                 f"60% от медианы: товар интересен, но на финише что-то мешает — цена в "
                 f"корзине, срок/стоимость доставки, наличие, конкурент дешевле. "
                 f"🟡 «карточка не убеждает» = мало кладут в корзину (контент/фото/цена). "
                 f"«Потеряно ₽» = сколько бы принесли недостающие заказы, дотяни корзина→заказ "
                 f"до медианы, по ставке ВЗВЕШ ₽/шт с листа «Сверка выплат» (только для SKU "
                 f"с положительной прибылью; иначе прочерк — доводить убыточное незачем). "
                 f"ВАЖНО: часть корзин — это «отложил на потом», а не отказ; поэтому цифра — "
                 f"верхняя граница потенциала, а не гарантированные деньги.")
        kheads = ["Товар", "SKU", "Клики", "Корзины", "Заказы",
                  "Клик→\nкорзина %", "Клик→\nзаказ %", "Корзина→\nзаказ %",
                  "Корзин без\nзаказа, шт", "Недобрано\nзаказов", "Потеряно ₽", "Диагноз"]
        nk = len(kheads)
        wsk.merge_cells(start_row=1, start_column=1, end_row=1, end_column=nk)
        wsk.cell(row=1, column=1, value=knote).font = _tfont(sz=10, color="555555")
        wsk.cell(row=1, column=1).alignment = _left()
        wsk.row_dimensions[1].height = 74
        for c, h in enumerate(kheads, 1):
            wsk.cell(row=2, column=c, value=h)
        _style_header_row(wsk, 2, nk, "6C3483")
        wsk.row_dimensions[2].height = 34

        rk = 2
        for x in rows_f:
            rk += 1
            vals = [name_by_sku.get(x["sku"], x["sku"])[:45], x["sku"],
                    int(x["clicks"]), int(x["carts"]), int(x["orders"]),
                    round(x["k2c"], 1), round(x["k2o"], 1),
                    round(x["c2o"], 1) if x["c2o"] is not None else "—",
                    int(x["no_order_carts"]), round(x["lost_orders"], 1),
                    x["lost_rub"] if x["lost_rub"] is not None else "—", x["diag"]]
            for c, v in enumerate(vals, 1):
                cell = wsk.cell(row=rk, column=c, value=v)
                cell.border = _border()
                cell.alignment = _left() if c in (1, nk) else _center() if c > 2 else _left()
                if c in (6, 7, 8) and isinstance(v, (int, float)):
                    cell.number_format = '0.0"%"'
                if c == 11 and isinstance(v, (int, float)):
                    cell.number_format = '#,##0'
            c2o_c = wsk.cell(row=rk, column=8)
            if x["c2o"] is not None and x["carts"] >= 3:
                if x["c2o"] < med_c2o * 0.6:
                    c2o_c.fill = _fill(RED_LIGHT); c2o_c.font = _tfont(bold=True, color="B71C1C")
                elif x["c2o"] >= med_c2o:
                    c2o_c.fill = _fill(GREEN_LT)
            k2c_c = wsk.cell(row=rk, column=6)
            if x["k2c"] < med_k2c * 0.6 and x["clicks"] >= 15:
                k2c_c.fill = _fill(YELLOW_LT); k2c_c.font = _tfont(bold=True, color="9A7D0A")
            if x["lost_rub"]:
                lc = wsk.cell(row=rk, column=11)
                lc.font = _tfont(bold=True, color="B71C1C")
        wsk.column_dimensions["A"].width = 38
        wsk.column_dimensions["B"].width = 14
        for i in range(3, nk):
            wsk.column_dimensions[get_column_letter(i)].width = 11
        wsk.column_dimensions[get_column_letter(nk)].width = 26
        wsk.freeze_panes = "C3"
        wsk.auto_filter.ref = f"A2:{get_column_letter(nk)}{rk}"

    # ══════════════════════════════════════════════════════════════════════════
    # ЛИСТ: ПРИБЫЛЬ ПО МЕСЯЦАМ (помесячная раскладка по SKU; та же формула, что
    # и Σ 90д на «Сверке выплат» — сырой агрегат из netting, профит здесь же)
    # ══════════════════════════════════════════════════════════════════════════
    netting_months = kwargs.get("netting_months", {}) or {}
    if netting_months:
        from .cost_prices import get_cost_price as _gcp_m

        month_keys = sorted(netting_months)  # "YYYY-MM" сортируется лексикографически = хронологически
        # агрегат по (SKU, месяц), сумма моделей FBS+FBY
        by_sku_month = {}   # sku -> {month_key: {qty, ret_qty, buyer_pay, payout, returns}}
        for mk in month_keys:
            for (sku, _model), a in netting_months[mk].items():
                acc = by_sku_month.setdefault(sku, {}).setdefault(
                    mk, {"qty": 0.0, "ret_qty": 0.0, "buyer_pay": 0.0, "payout": 0.0, "returns": 0.0})
                for k in ("qty", "ret_qty", "buyer_pay", "payout", "returns"):
                    acc[k] += a.get(k, 0)

        def _month_profit(agg):
            if not agg or agg.get("qty", 0) <= 0:
                return None  # нет продаж в этом месяце — не 0, а «нет данных»
            net_q = agg["qty"] - agg.get("ret_qty", 0)
            if net_q <= 0:
                return 0.0
            cost = (_gcp_m(sku_row) or 0) * net_q
            tax = (agg["buyer_pay"] + agg.get("returns", 0)) * 0.07
            return round(agg["payout"] - cost - tax, 2)

        def _month_net_q(agg):
            """Чистые продажи = продано − возвраты (на них и делится прибыль)."""
            if not agg or agg.get("qty", 0) <= 0:
                return None
            return agg["qty"] - agg.get("ret_qty", 0)

        # прибыль по SKU по месяцам + итог по SKU (для сортировки топ-вниз)
        sku_profits = {}      # sku -> {month_key: profit|None}
        sku_netq = {}         # sku -> {month_key: чистые штуки|None}
        sku_totals = {}       # sku -> Σ прибыль по всем месяцам
        for sku_row, months_agg in by_sku_month.items():
            row = {}
            qrow = {}
            for mk in month_keys:
                row[mk] = _month_profit(months_agg.get(mk))
                qrow[mk] = _month_net_q(months_agg.get(mk))
            sku_profits[sku_row] = row
            sku_netq[sku_row] = qrow
            sku_totals[sku_row] = sum(v for v in row.values() if v is not None)

        grand_total = sum(sku_totals.values())
        month_totals = {mk: sum(v.get(mk) or 0 for v in sku_profits.values()) for mk in month_keys}
        month_netq = {mk: sum(v.get(mk) or 0 for v in sku_netq.values()) for mk in month_keys}

        wsm = wb.create_sheet("Прибыль по месяцам")
        wsm.sheet_properties.tabColor = "8E44AD"
        mnote = ("Помесячная чистая прибыль ФАКТ по SKU (та же формула, что «Σ 90д» на «Сверке выплат»): "
                 "прибыль = К перечисл. (united-netting) − себест.(чистые штуки) − налог 7% (с нетто-выручки). "
                 "Завершённые месяцы кэшируются навсегда (не пересчитываются задним числом); последний "
                 "столбец — текущий незавершённый месяц, обновляется каждый прогон. Пустая ячейка — "
                 "продаж по SKU в этом месяце не было (не 0). Отсортировано по вкладу в общую прибыль ↓. "
                 "На каждый месяц три колонки: «шт. чистые» — продано минус возвраты (именно на них "
                 "делится прибыль: возврат уезжает обратно на склад, себестоимость по нему не списывается), "
                 "«прибыль ₽» и «₽/шт». В строке ИТОГО ₽/шт — это Σприбыли ÷ Σштук по каталогу, "
                 "а не среднее по колонке. Цвет ₽/шт как на «Сверке выплат»: >100 ₽ зелёный, "
                 "50–100 ₽ жёлтый, <50 ₽ красный. После «Итого ₽» — «Итого шт.» (чистые штуки "
                 "за все месяцы листа) и «Итого ₽/шт» = Итого ₽ ÷ Итого шт.: прибыль на штуку за "
                 "весь период, не среднее помесячных ₽/шт.")
        wsm.merge_cells(f"A1:{get_column_letter(2 + 3 * len(month_keys) + 4)}1")
        wsm["A1"] = mnote
        wsm["A1"].font = _tfont(sz=10, color="555555"); wsm["A1"].alignment = _left()
        wsm.row_dimensions[1].height = 42

        # На каждый месяц — три колонки: чистые продажи (шт), прибыль ₽ и ₽/шт.
        # Шапка двухэтажная: месяц объединён над своей тройкой, под ним — подписи.
        MONTH_COL0 = 3          # первая колонка первого месяца
        PER_MONTH = 3
        # Свой оттенок на каждый ТИП колонки (штуки / деньги / деньги на штуку) —
        # месяцев много, и без этого взгляд теряет, в каком столбце что.
        TINT_QTY, TINT_RUB, TINT_PPU = "E8F1F8", "FDF3E2", "EDE7F6"
        TINTS = (TINT_QTY, TINT_RUB, TINT_PPU)
        TOTAL_COL = MONTH_COL0 + PER_MONTH * len(month_keys)
        TOTAL_QTY_COL = TOTAL_COL + 1   # чистые штуки за все месяцы листа
        TOTAL_PPU_COL = TOTAL_COL + 2   # Итого ₽ ÷ Итого шт.
        SHARE_COL = TOTAL_COL + 3
        hr_m = 2                # строка с месяцами
        hr_sub = 3              # строка с подписями колонок

        wsm.cell(row=hr_m, column=1, value="Товар")
        wsm.cell(row=hr_m, column=2, value="SKU")
        wsm.merge_cells(start_row=hr_m, start_column=1, end_row=hr_sub, end_column=1)
        wsm.merge_cells(start_row=hr_m, start_column=2, end_row=hr_sub, end_column=2)
        for i, mk in enumerate(month_keys):
            base = MONTH_COL0 + i * PER_MONTH
            wsm.cell(row=hr_m, column=base, value=mk)
            wsm.merge_cells(start_row=hr_m, start_column=base,
                            end_row=hr_m, end_column=base + PER_MONTH - 1)
            for off, sub in enumerate(("шт.\nчистые", "прибыль ₽", "₽/шт")):
                wsm.cell(row=hr_sub, column=base + off, value=sub)
        for col, title in ((TOTAL_COL, "Итого ₽"), (TOTAL_QTY_COL, "Итого\nшт."),
                           (TOTAL_PPU_COL, "Итого\n₽/шт"), (SHARE_COL, "Доля %")):
            wsm.cell(row=hr_m, column=col, value=title)
            wsm.merge_cells(start_row=hr_m, start_column=col, end_row=hr_sub, end_column=col)
        for hrow in (hr_m, hr_sub):
            _style_header_row(wsm, hrow, SHARE_COL, "8E44AD")
        # подписи под месяцем красим в цвет своего типа — шапка работает легендой
        for i in range(len(month_keys)):
            base = MONTH_COL0 + i * PER_MONTH
            for off, tint in enumerate(TINTS):
                sc_h = wsm.cell(row=hr_sub, column=base + off)
                sc_h.fill = _fill(tint)
                sc_h.font = _tfont(bold=True, color="333333")
        wsm.row_dimensions[hr_m].height = 18
        wsm.row_dimensions[hr_sub].height = 26

        def _ppu_font(v, bold=False):
            """₽/шт теми же порогами, что на «Сверке выплат»: >100 / 50–100 / <50.
            Здесь только цветом шрифта — фоном на пять месяцев получилась бы рябь."""
            return _tfont(bold=bold, color="1B5E20" if v > 100 else
                          "7D6608" if v >= 50 else "B71C1C")

        r = hr_sub
        for sku_row in sorted(sku_totals, key=lambda s: -sku_totals[s]):
            r += 1
            share = round(sku_totals[sku_row] / grand_total * 100, 1) if grand_total else 0
            wsm.cell(row=r, column=1, value=name_by_sku.get(sku_row, sku_row)[:50])
            wsm.cell(row=r, column=2, value=sku_row)
            for i, mk in enumerate(month_keys):
                base = MONTH_COL0 + i * PER_MONTH
                v = sku_profits[sku_row][mk]
                nq = sku_netq[sku_row][mk]
                qc = wsm.cell(row=r, column=base, value=int(nq) if nq is not None else None)
                qc.fill = _fill(TINT_QTY)
                if nq is not None:
                    qc.number_format = '#,##0'
                cell = wsm.cell(row=r, column=base + 1, value=v)
                cell.fill = _fill(TINT_RUB)
                if v is not None:
                    cell.number_format = '#,##0'
                    cell.font = _tfont(color="1B5E20" if v > 0 else "B71C1C" if v < 0 else "999999")
                # ₽/шт считаем только когда есть чистые штуки: иначе делить не на что
                ppu = round(v / nq, 2) if (v is not None and nq) else None
                pc = wsm.cell(row=r, column=base + 2, value=ppu)
                pc.fill = _fill(TINT_PPU)
                if ppu is not None:
                    pc.number_format = '#,##0'
                    pc.font = _ppu_font(ppu)
            tc = wsm.cell(row=r, column=TOTAL_COL, value=round(sku_totals[sku_row]))
            tc.number_format = '#,##0'
            tc.font = _tfont(bold=True, color="1B5E20" if sku_totals[sku_row] > 0 else "B71C1C")
            # штуки берём только за месяцы, где есть прибыль, — те же, что сложены в Итого ₽
            tot_q = sum(sku_netq[sku_row][mk] or 0 for mk in month_keys
                        if sku_profits[sku_row][mk] is not None)
            qc = wsm.cell(row=r, column=TOTAL_QTY_COL, value=int(tot_q))
            qc.number_format = '#,##0'; qc.fill = _fill(TINT_QTY)
            tot_ppu = round(sku_totals[sku_row] / tot_q, 2) if tot_q else None
            pc = wsm.cell(row=r, column=TOTAL_PPU_COL, value=tot_ppu)
            pc.fill = _fill(TINT_PPU)
            if tot_ppu is not None:
                pc.number_format = '#,##0'
                pc.font = _ppu_font(tot_ppu, bold=True)
            sc = wsm.cell(row=r, column=SHARE_COL, value=share)
            if share >= 10:
                sc.fill = _fill(GREEN_LT)
            elif share >= 5:
                sc.fill = _fill(YELLOW_LT)
            for c in range(1, SHARE_COL + 1):
                cell = wsm.cell(row=r, column=c)
                cell.border = _border()
                cell.alignment = _left() if c == 1 else _center()

        # ИТОГО по месяцам
        r += 1
        wsm.cell(row=r, column=1, value="ИТОГО")
        for i, mk in enumerate(month_keys):
            base = MONTH_COL0 + i * PER_MONTH
            qc = wsm.cell(row=r, column=base, value=int(month_netq[mk]))
            qc.number_format = '#,##0'
            cell = wsm.cell(row=r, column=base + 1, value=round(month_totals[mk]))
            cell.number_format = '#,##0'
            # ₽/шт по каталогу за месяц — не среднее из колонки, а Σприбыль/Σштуки
            ppu = round(month_totals[mk] / month_netq[mk], 2) if month_netq[mk] else None
            pc = wsm.cell(row=r, column=base + 2, value=ppu)
            if ppu is not None:
                pc.number_format = '#,##0'
        tc = wsm.cell(row=r, column=TOTAL_COL, value=round(grand_total))
        tc.number_format = '#,##0'
        grand_q = sum(month_netq.values())
        wsm.cell(row=r, column=TOTAL_QTY_COL, value=int(grand_q)).number_format = '#,##0'
        grand_ppu = round(grand_total / grand_q, 2) if grand_q else None
        wsm.cell(row=r, column=TOTAL_PPU_COL, value=grand_ppu).number_format = '#,##0'
        for c in range(1, SHARE_COL + 1):
            cell = wsm.cell(row=r, column=c)
            cell.border = _border(); cell.fill = _fill(GREY_LT); cell.font = _tfont(bold=True)
            cell.alignment = _left() if c == 1 else _center()
        if grand_ppu is not None:
            wsm.cell(row=r, column=TOTAL_PPU_COL).font = _ppu_font(grand_ppu, bold=True)
        for i, mk in enumerate(month_keys):   # цвет ₽/шт по тем же порогам
            ppu = wsm.cell(row=r, column=MONTH_COL0 + i * PER_MONTH + 2).value
            if isinstance(ppu, (int, float)):
                wsm.cell(row=r, column=MONTH_COL0 + i * PER_MONTH + 2).font = _ppu_font(ppu, bold=True)

        # MoM % — прирост прибыли к предыдущему месяцу (в колонке прибыли)
        r += 1
        wsm.cell(row=r, column=1, value="MoM %")
        for i, mk in enumerate(month_keys):
            if i == 0:
                continue
            prev_v = month_totals[month_keys[i - 1]]
            cur_v = month_totals[mk]
            pct = round((cur_v - prev_v) / prev_v * 100, 1) if prev_v else None
            cell = wsm.cell(row=r, column=MONTH_COL0 + i * PER_MONTH + 1,
                            value=(f"{'+' if pct >= 0 else ''}{pct}%" if pct is not None else "—"))
            cell.font = _tfont(bold=True, color="1B5E20" if (pct or 0) >= 0 else "B71C1C")
        for c in range(1, SHARE_COL + 1):
            cell = wsm.cell(row=r, column=c)
            cell.border = _border(); cell.alignment = _left() if c == 1 else _center()
            if c == 1:
                cell.font = _tfont(bold=True)

        wsm.column_dimensions["A"].width = 40
        wsm.column_dimensions["B"].width = 13
        for i, mk in enumerate(month_keys):
            base = MONTH_COL0 + i * PER_MONTH
            wsm.column_dimensions[get_column_letter(base)].width = 7      # шт
            wsm.column_dimensions[get_column_letter(base + 1)].width = 11  # ₽
            wsm.column_dimensions[get_column_letter(base + 2)].width = 8   # ₽/шт
        wsm.column_dimensions[get_column_letter(TOTAL_COL)].width = 12
        wsm.column_dimensions[get_column_letter(TOTAL_QTY_COL)].width = 8
        wsm.column_dimensions[get_column_letter(TOTAL_PPU_COL)].width = 9
        wsm.column_dimensions[get_column_letter(SHARE_COL)].width = 9
        wsm.freeze_panes = f"C{hr_sub + 1}"
        wsm.auto_filter.ref = f"A{hr_sub}:{get_column_letter(SHARE_COL)}{r-2}"

    # ══════════════════════════════════════════════════════════════════════════
    # ПОРЯДОК ЛИСТОВ
    # ══════════════════════════════════════════════════════════════════════════
    desired_order = [
        "Дашборд", "Скидки FEE по заказам", "Сверка выплат", "Прибыль по месяцам",
        "Юнит-экономика",
        # Листы динамики идут в порядке ВОРОНКИ: показы → клики → CTR →
        # корзины → клик/корзина → покупки → клик/покупка.
        "Товары", "FBY остатки по дням", "Показы 14д",
        "Клики 14д", "CTR 14д", "Корзины 14д", "CR клик→корзина 14д",
        "Покупки 14д", "CR клик→покупка 14д",
        "Корзина vs Покупка", "Конверсия карточек", "Дырявая воронка",
        "Индекс цен", "Видимость↔Буст↔Показы",
        "FBS FBY Остатки", "Ставки буста по дням",
    ]
    trailing_order = ["Качество и отзывы"]
    ordered_sheets = [wb[name] for name in desired_order if name in wb.sheetnames]
    ordered_sheets += [wb[name] for name in wb.sheetnames
                        if name not in desired_order and name not in trailing_order]
    ordered_sheets += [wb[name] for name in trailing_order if name in wb.sheetnames]
    wb._sheets = ordered_sheets

    # ══════════════════════════════════════════════════════════════════════════
    # SAVE
    # ══════════════════════════════════════════════════════════════════════════
    fname = os.path.join(report_folder, f"ym_report_{run_date.isoformat()}.xlsx")
    wb.save(fname)
    return fname
