"""
Yandex Market Partner API client.

Auth: Api-Key header (ACMA:... format token)

Key endpoints:
  POST /businesses/{businessId}/offer-cards          — list of all offers
  POST /reports/shows-sales/generate                 — analytics report (shows/clicks/cart/orders/boost)
  GET  /reports/info/{reportId}                      — poll report status → get download URL

Report columns (shows-sales, grouping=OFFERS):
  День, Категория, Бренд, Ваш SKU, Название товара,
  Индекс видимости %,
  Показы моих товаров шт,  Показы моих товаров с акциями шт,  Доля показов с бустом %,
  Клики по товарам шт,     Клики по товарам с акциями шт,
  Конверсия из показа в корзину %,
  Добавления в корзину шт, Добавления в корзину по акциям шт, Доля добавлений с бустом %,
  Заказанные товары шт,    Заказанные товары по акциям шт,
  Заказано на сумму ₽,     Заказано с акциями на сумму ₽,
  Конверсия из корзины в заказ %,
  Доля заказанных товаров с бустом %,
  ...
"""

import time
import io
import requests
import openpyxl
import pandas as pd
from datetime import date, datetime, timedelta
from .config import OAUTH_TOKEN, BUSINESS_ID, YM_API_BASE


def _headers():
    return {"Api-Key": OAUTH_TOKEN, "Content-Type": "application/json"}


# Кратковременные сбои сети/DNS (например, не разрешился api.partner.market.yandex.ru)
# не должны ронять весь прогон — ретраим с нарастающей паузой.
_NET_ERRORS = (requests.exceptions.ConnectionError, requests.exceptions.Timeout)


def _request_raw_with_retry(method, url, *, retries=5, timeout=60, **kwargs):
    """requests.request с ретраем сетевых сбоев (DNS/connection/timeout).
    Возвращает «сырой» Response — без raise_for_status, чтобы вызывающий
    код мог сам обработать спец-статусы (напр. 420 rate limit)."""
    for attempt in range(retries):
        try:
            r = requests.request(method, url, timeout=timeout, **kwargs)
        except _NET_ERRORS as e:
            if attempt == retries - 1:
                raise
            wait = 10 * (attempt + 1)
            print(f"  сетевой сбой на {url} ({type(e).__name__}), "
                  f"повтор через {wait}s ({attempt+1}/{retries})...")
            time.sleep(wait)
            continue
        # 5xx — сбой на стороне Яндекса, почти всегда кратковременный. Раньше он
        # не ретраился и одна случайная 500 на первом же /offer-cards роняла весь
        # прогон (20.08.2026 отчёт не собрался именно так). 4xx НЕ ретраим —
        # это наши ошибки (401/400), сами не пройдут.
        if r.status_code >= 500 and attempt < retries - 1:
            wait = 10 * (attempt + 1)
            print(f"  {r.status_code} от сервера на {url}, "
                  f"повтор через {wait}s ({attempt+1}/{retries})...")
            time.sleep(wait)
            continue
        return r


def _request_with_retry(method, path, *, retries=5, **kwargs):
    url = f"{YM_API_BASE}{path}"
    r = _request_raw_with_retry(method, url, headers=_headers(), retries=retries, **kwargs)
    r.raise_for_status()
    return r.json()


def _get(path, params=None, retries=5):
    return _request_with_retry("GET", path, params=params, retries=retries)


def _post(path, body, params=None, retries=5):
    return _request_with_retry("POST", path, json=body, params=params, retries=retries)


# ── Offer catalogue ──────────────────────────────────────────────────────────

def get_all_offers() -> list[dict]:
    """
    Returns list of {shopSku, marketSku, name} for every offer in the catalogue.
    Uses POST /businesses/{id}/offer-cards with pagination.
    NOTE: pageToken must be passed as a query parameter, not in the request body.
    """
    offers = []
    page_token = None
    seen_tokens = set()
    while True:
        params = {"pageToken": page_token} if page_token else {}
        data = _post(f"/businesses/{BUSINESS_ID}/offer-cards", {}, params=params)
        for o in data.get("result", {}).get("offerCards", []):
            mapping = o.get("mapping", {})
            offers.append({
                "shopSku":   o.get("offerId", ""),
                "marketSku": mapping.get("marketSku"),
                "name":      o.get("basicOffer", {}).get("name", o.get("offerId", "")),
                "content_rating": o.get("contentRating"),     # 0–100, качество карточки
                "card_status":    o.get("cardStatus", ""),    # напр. HAS_CARD_CAN_UPDATE
            })
        page_token = data.get("result", {}).get("paging", {}).get("nextPageToken")
        if not page_token or page_token in seen_tokens:
            break
        seen_tokens.add(page_token)
        time.sleep(0.2)
    return offers


# ── Shows-Sales report ────────────────────────────────────────────────────────

# Column name mapping from Russian report headers to short internal keys
_COL_MAP = {
    "Ваш SKU":                                    "shopSku",
    "Название товара":                            "name",
    "Категория":                                  "category",
    "Индекс видимости, %":                        "visibility",
    "Показы моих товаров, шт.":                   "shows",
    "Показы моих товаров с акциями, шт.":         "shows_promo",
    "Доля показов с бустом, %":                   "shows_boost_pct",
    "Клики по товарам, шт.":                      "clicks",
    "Клики по товарам с акциями, шт.":            "clicks_promo",
    "Конверсия из показа в корзину, %":           "ctr_to_cart",
    "Добавления в корзину, шт.":                  "carts",
    "Добавления в корзину по акциям, шт.":        "carts_promo",
    "Доля добавлений товаров с бустом в корзину, %": "carts_boost_pct",
    "Заказанные товары, шт.":                     "orders",
    "Заказанные товары по акциям, шт.":           "orders_promo",
    "Заказано товаров на сумму, ₽":               "revenue",
    "Заказано товаров с акциями на сумму, ₽":     "revenue_promo",
    "Конверсия из корзины в заказ, %":            "cart_to_order",
    "Доля заказанных товаров с бустом, %":        "orders_boost_pct",
}


def _generate_shows_sales_report(date_from: date, date_to: date) -> str:
    """
    Trigger report generation. Returns reportId.
    Rate limit: 1 request per 10 minutes per businessId.
    Auto-retries on 420 (rate limit) up to 3 times with 10-min waits.
    """
    body = {
        "businessId": int(BUSINESS_ID),
        "dateFrom":   date_from.strftime("%Y-%m-%d"),
        "dateTo":     date_to.strftime("%Y-%m-%d"),
        "grouping":   "OFFERS",
    }
    for attempt in range(3):
        r = _request_raw_with_retry(
            "POST", f"{YM_API_BASE}/reports/shows-sales/generate",
            headers=_headers(), json=body,
        )
        if r.status_code == 420:
            wait = 620  # 10 min + 20s buffer
            print(f"  Rate limit hit (attempt {attempt+1}/3), waiting {wait}s...")
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r.json()["result"]["reportId"]
    raise RuntimeError("shows-sales rate limit: 3 retries exhausted")


def _wait_for_report(report_id: str, max_wait_sec: int = 120) -> str:
    """Poll /reports/info/{id} until DONE, return download URL."""
    for _ in range(max_wait_sec // 5):
        time.sleep(5)
        data = _get(f"/reports/info/{report_id}")
        result = data.get("result", {})
        status = result.get("status")
        if status == "DONE":
            return result["file"]
        if status in ("ERROR", "FAILED"):
            raise RuntimeError(f"Report {report_id} failed: {result}")
    raise TimeoutError(f"Report {report_id} not ready after {max_wait_sec}s")


def _download_report_df(file_url: str) -> pd.DataFrame:
    """Download xlsx and return as DataFrame with renamed columns."""
    r = _request_raw_with_retry("GET", file_url)
    r.raise_for_status()
    df = pd.read_excel(io.BytesIO(r.content))
    # Rename known columns
    df = df.rename(columns={k: v for k, v in _COL_MAP.items() if k in df.columns})
    return df


def fetch_shows_sales(date_from: date, date_to: date) -> pd.DataFrame:
    """
    Generate and download the shows-sales analytics report.
    Returns DataFrame with one row per (day × shopSku).

    ⚠️  Rate limit: 1 report per 10 min. The agent generates ONE report
        covering the full needed date range (today - 7d to today),
        then slices the DataFrame as needed.
    """
    print(f"  Generating shows-sales report {date_from} → {date_to}...")
    report_id = _generate_shows_sales_report(date_from, date_to)
    print(f"  reportId={report_id}, waiting...")
    file_url = _wait_for_report(report_id)
    df = _download_report_df(file_url)
    print(f"  Report ready: {len(df)} rows, {len(df.columns)} columns")
    return df


# ── United Marketplace Services report (комиссии, в т.ч. скидки за акции) ───

def _generate_ums_report(date_from: date, date_to: date) -> str:
    """
    Trigger united-marketplace-services report generation. Returns reportId.
    Rate limit: 1 request per 10 minutes per businessId.
    Auto-retries on 420 up to 3 times with 10-min waits.
    """
    body = {
        "businessId": int(BUSINESS_ID),
        "dateFrom":   date_from.strftime("%Y-%m-%d"),
        "dateTo":     date_to.strftime("%Y-%m-%d"),
    }
    for attempt in range(3):
        r = _request_raw_with_retry(
            "POST", f"{YM_API_BASE}/reports/united-marketplace-services/generate",
            headers=_headers(), json=body,
        )
        if r.status_code == 420:
            wait = 620
            print(f"  Rate limit hit (attempt {attempt+1}/3), waiting {wait}s...")
            time.sleep(wait)
            continue
        r.raise_for_status()
        return r.json()["result"]["reportId"]
    raise RuntimeError("united-marketplace-services rate limit: 3 retries exhausted")


def fetch_placement_fee_detail(date_from: date, date_to: date) -> list[dict]:
    """
    Generate and download the united-marketplace-services report, and return
    per-order/per-SKU rows from sheet «Размещение товаров и услуг» (Яндекс
    переименовал лист с «...на витрине» ~02.07.2026; колонки не изменились) with the
    full (pre-discount) commission, the joint-promo discount, and the final
    amount actually charged for placement.

    Date range is filtered by Yandex on «Дата и время оказания услуги» (AH).

    Returns list of dicts:
        {date, order_id, sku, name, price, qty, fee_full, fee_discount, fee_paid}
    """
    print(f"  Generating united-marketplace-services report {date_from} → {date_to}...")
    report_id = _generate_ums_report(date_from, date_to)
    print(f"  reportId={report_id}, waiting...")
    file_url = _wait_for_report(report_id, max_wait_sec=180)

    r = _request_raw_with_retry("GET", file_url, retries=5, timeout=120)
    r.raise_for_status()
    wb = openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)
    ws = wb["Размещение товаров и услуг"]

    rows = []
    for row in ws.iter_rows(min_row=7, values_only=False):
        sku = row[9].value  # J
        if not sku:
            continue
        fee_full = row[35].value or 0      # AJ
        fee_discount = row[48].value or 0  # AW
        fee_paid = row[49].value or 0      # AX
        fee_full = round(float(fee_full), 2)
        fee_discount = round(float(fee_discount), 2)
        fee_paid = round(float(fee_paid), 2)
        # AX = AJ - AL + AR + AT - AU - AV - AW, т.е. кроме скидки за акцию (AW)
        # на итог влияют штрафы за просрочку (AR), скидка за лояльность (AV) и др.
        # other_adj собирает суммарно все эти прочие корректировки.
        other_adj = round(fee_paid - (fee_full - fee_discount), 2)
        rows.append({
            "date":     str(row[33].value)[:10] if row[33].value else "",  # AH
            "order_id": row[7].value,   # H
            "sku":      sku,            # J
            "name":     row[10].value,  # K
            "price":    row[12].value or 0,  # M
            "qty":      row[14].value or 0,  # O
            "fee_full":     fee_full,
            "fee_discount": fee_discount,
            "fee_paid":     fee_paid,
            "other_adj":    other_adj,
        })

    # ── Лист «Буст продаж, оплата за продажи» — РЕАЛЬНЫЙ буст по SKU ──
    # Колонки (1-based): I(9)=SKU, M(13)=кол-во, O(15)=сработавшая ставка %,
    # Q(17)=постоплата ₽ (gross буст), S(19)=скидка за совм.акции.
    # Net буст = постоплата − скидка. Сработавшая ставка обычно << потолка.
    boost_by_sku = {}
    if "Буст продаж, оплата за продажи" in wb.sheetnames:
        wsb = wb["Буст продаж, оплата за продажи"]
        for row in wsb.iter_rows(min_row=3, values_only=True):
            sku = row[8] if len(row) > 8 else None  # I
            if not sku:
                continue
            qty = float(row[12] or 0)           # M
            rate = float(row[14] or 0)          # O — сработавшая ставка %
            post = float(row[16] or 0)          # Q — постоплата (gross)
            disc = float(row[18] or 0) if len(row) > 18 else 0.0  # S — скидка совм.акций
            a = boost_by_sku.setdefault(sku, {"qty": 0.0, "gross": 0.0, "net": 0.0,
                                              "rate_w": 0.0})
            a["qty"]   += qty
            a["gross"] += post
            a["net"]   += post - disc
            a["rate_w"] += rate * qty   # для взвешенного среднего по кол-ву
    for a in boost_by_sku.values():
        a["rate_avg"] = round(a["rate_w"] / a["qty"], 2) if a["qty"] else 0.0
        a["gross_per_unit"] = round(a["gross"] / a["qty"], 2) if a["qty"] else 0.0
        a["net_per_unit"] = round(a["net"] / a["qty"], 2) if a["qty"] else 0.0

    print(f"  Report ready: {len(rows)} fee rows, boost SKU: {len(boost_by_sku)}")
    return rows, boost_by_sku


# ── Ставки буста продаж ────────────────────────────────────────────────────────

def get_bid_recommendations(skus: list[str]) -> dict:
    """Рекомендованные ставки буста по SKU.
    POST /businesses/{id}/bids/recommendations. Батч до 1500 SKU.

    Ставка в API = % от цены × 100 (1010 = 10.1%). Кривая bidRecommendations
    показывает, какая ставка нужна для какой доли показов (showPercent).

    Returns: {sku: {"bid": int|None, "curve": [(bid_pct, show_pct), ...]}}
      bid — рекомендованная (для макс. показов), bid_pct = bid/100.
    """
    out = {}
    for i in range(0, len(skus), 1500):
        batch = skus[i:i + 1500]
        # best-effort фаза: при сбое сети скипаем быстро (retries=2), не крутим 5 ретраев
        data = _post(f"/businesses/{BUSINESS_ID}/bids/recommendations", {"skus": batch},
                     retries=2)
        for x in data.get("result", {}).get("recommendations", []):
            sku = x.get("sku")
            if not sku:
                continue
            curve = [(round(c.get("bid", 0) / 100, 2), c.get("showPercent"))
                     for c in x.get("bidRecommendations", [])]
            out[sku] = {
                "bid": x.get("bid"),
                "bid_pct": round(x["bid"] / 100, 2) if x.get("bid") is not None else None,
                "curve": curve,
            }
        time.sleep(0.2)
    return out


def get_current_bids(skus: list[str]) -> dict:
    """Текущие ставки буста, выставленные через API.
    POST /businesses/{id}/bids/info. ВНИМАНИЕ: при управлении бустом через
    «ценовую стратегию»/кампании кабинета (а не per-SKU API) возвращает пусто.

    Returns: {sku: bid_pct}
    """
    out = {}
    data = _post(f"/businesses/{BUSINESS_ID}/bids/info", {"skus": skus})
    for x in data.get("result", {}).get("bids", []):
        if x.get("sku") and x.get("bid") is not None:
            out[x["sku"]] = round(x["bid"] / 100, 2)
    return out


def set_prices(prices: dict[str, float], currency: str = "RUR", batch_size: int = 500) -> int:
    """Выставляет базовую цену продавца per-SKU. POST /businesses/{id}/offer-prices/updates.

    prices: {sku: цена в рублях} (или иной валюте через currency).
    Это ЦЕНА ПРОДАВЦА (basicPrice) — та же, что читает get_catalog_prices(). Это
    НЕ промо-цена и не цена на витрине с учётом акций/баллов — те считаются от
    этой базовой отдельно.

    Returns: сколько SKU отправлено.
    """
    items = [{"offerId": sku, "price": {"value": price, "currencyId": currency}}
             for sku, price in prices.items()]
    sent = 0
    for i in range(0, len(items), batch_size):
        batch = items[i:i + batch_size]
        _request_with_retry("POST", f"/businesses/{BUSINESS_ID}/offer-prices/updates",
                            json={"offers": batch})
        sent += len(batch)
        time.sleep(0.3)
    return sent


def set_bids(bids: dict[str, float], batch_size: int = 500) -> int:
    """Выставляет ставки буста per-SKU. PUT /businesses/{id}/bids.

    bids: {sku: процент} — ИМЕННО процент (20.1), не «сырое» значение API.
    В API bid = % × 100 (20.1% → 2010), конвертация здесь, чтобы наружу
    везде ходили человеческие проценты.

    ⚠️ bid=0 НЕ выключает буст: это снимает per-SKU override, и товар падает
    под групповую кампанию кабинета, наследуя её ставку. Для гарантированного
    «почти выкл» ставь 1 (=1%), а не 0.

    Ставка — потолок: аукцион второй цены обычно списывает МЕНЬШЕ выставленного.

    Returns: сколько SKU отправлено.
    """
    items = [{"sku": sku, "bid": int(round(pct * 100))} for sku, pct in bids.items()]
    sent = 0
    for i in range(0, len(items), batch_size):
        batch = items[i:i + batch_size]
        _request_with_retry("PUT", f"/businesses/{BUSINESS_ID}/bids",
                            json={"bids": batch})
        sent += len(batch)
        time.sleep(0.3)
    return sent


# ── Отчёт по платежам (united-netting) — РЕАЛЬНЫЕ выплаты по артикулам ──────────

def _generate_netting_report(date_from: date, date_to: date) -> str:
    """Trigger united-netting report generation. Returns reportId.
    Лимиты как у др. отчётов; ретрай на 420."""
    body = {"businessId": int(BUSINESS_ID),
            "dateFrom": date_from.strftime("%Y-%m-%d"),
            "dateTo": date_to.strftime("%Y-%m-%d")}
    for attempt in range(3):
        r = _request_raw_with_retry(
            "POST", f"{YM_API_BASE}/reports/united-netting/generate",
            headers=_headers(), json=body)
        if r.status_code == 420:
            print(f"  netting rate limit ({attempt+1}/3), ждём 130s...")
            time.sleep(130)
            continue
        r.raise_for_status()
        return r.json()["result"]["reportId"]
    raise RuntimeError("united-netting rate limit exhausted")


_NETTING_HEADERS = {
    "model":  "модели работы",
    "date":   "дата транзакции",
    "sku":    "ваш sku",
    "qty":    "количество",
    "amount": "сумма транзакции",
    "source": "источник транзакции",
}


def _netting_columns(ws) -> dict:
    """{поле: 0-based индекс колонки} по заголовкам строки 2 отчёта «По платежам».
    Если хоть одного обязательного заголовка нет (или он неоднозначен) — падаем,
    чтобы вызывающий откатился на прошлый кэш, а не записал мусор поверх него."""
    header = [" ".join(str(c.value or "").split()).lower() for c in ws[2]]
    col = {}
    for field, needle in _NETTING_HEADERS.items():
        hits = [i for i, h in enumerate(header) if needle in h]
        if len(hits) != 1:
            raise RuntimeError(
                f"структура отчёта «По платежам» изменилась: заголовок «{needle}» "
                f"найден {len(hits)} раз (ожидался 1). Колонки: {header}")
        col[field] = hits[0]
    return col


def fetch_netting(date_from: date, date_to: date, recent_days: int | None = None,
                  strict_window: bool = False) -> dict:
    """РЕАЛЬНЫЕ движения денег от Яндекса по артикулам за период.
    Отчёт «По платежам» (united-netting), лист «Отчёт о платежах».

    Колонки ищутся ПО ЗАГОЛОВКАМ (строка 2), а не по позициям: Яндекс уже дважды
    менял структуру (03.09.2026 вставил "Ставка НДС"/"Сумма НДС, ₽" перед "Тип
    транзакции" — источник сдвинулся с 23-й колонки на 25-ю, и полтора дня весь
    netting молча падал в "Прочее" с Продано=0, успев закэшироваться).
    Сумма транзакций по SKU = чистая выплата продавцу (уже минус комиссии/буст/возвраты).

    Если recent_days задан — попутно (без доп. запроса) строится второй агрегат
    только по строкам с датой транзакции >= date_to - recent_days + 1 (напр. «последние
    7 дней» внутри уже загруженного 30-дневного отчёта).

    Returns: {(sku, model): {payout, buyer_pay, returns, services, other, qty}}
             либо (out, out_recent), если recent_days задан.
    """
    # Генерация отчёта у Яндекса иногда отваливается со статусом FAILED за доли
    # секунды — разовый сбой на их стороне (10.09.2026 так пропал весь столбец
    # текущего месяца на листе «Прибыль по месяцам»). Битый отчёт не чинится
    # опросом: нужен новый reportId, поэтому повторяем генерацию целиком.
    for attempt in range(3):
        print(f"  Generating united-netting report {date_from} → {date_to}...")
        rid = _generate_netting_report(date_from, date_to)
        print(f"  reportId={rid}, waiting...")
        try:
            url = _wait_for_report(rid, max_wait_sec=240)
            break
        except (RuntimeError, TimeoutError) as e:
            if attempt == 2:
                raise
            print(f"  отчёт не собрался ({str(e)[:90]}), перегенерируем через 30s "
                  f"({attempt + 1}/3)...")
            time.sleep(30)
    r = _request_raw_with_retry("GET", url, retries=5, timeout=180)
    r.raise_for_status()
    wb = openpyxl.load_workbook(io.BytesIO(r.content), data_only=True)
    ws = wb["Отчёт о платежах"]
    col = _netting_columns(ws)
    out = {}
    out_recent = {} if recent_days is not None else None
    recent_cutoff = (date_to - timedelta(days=recent_days - 1)) if recent_days else None
    unknown_sources = {}

    def _acc(bucket, key, amount, qty, s):
        a = bucket.setdefault(key, {"payout": 0.0, "buyer_pay": 0.0, "returns": 0.0,
                                    "services": 0.0, "promo": 0.0, "points": 0.0,
                                    "other": 0.0, "qty": 0.0, "ret_qty": 0.0})
        a["payout"] += amount
        if "платёж покупател" in s or "платеж покупател" in s:
            a["buyer_pay"] += amount
            a["qty"] += qty
        elif "возврат платеж" in s:
            a["returns"] += amount               # реальный возврат денег покупателю
            a["ret_qty"] += qty if qty else 1    # кол-во возвращённых штук
        elif "оплата услуг" in s or "удержан" in s:
            a["services"] += amount              # комиссия + буст + логистика + хранение
        elif "балл" in s:
            a["points"] += amount                # баллы-компенсация (+); возврат баллов (−)
        elif "скидк" in s:
            a["promo"] += amount                 # скидка за участие в акциях (−); возврат (+)
        else:
            a["other"] += amount                 # на всякий случай — не должно быть
            unknown_sources[s] = unknown_sources.get(s, 0.0) + amount

    need = max(col.values())
    for row in ws.iter_rows(min_row=3, values_only=True):
        if len(row) <= need:
            continue
        sku = row[col["sku"]]
        if not sku:
            continue
        model = (row[col["model"]] or "").upper()
        qty = float(row[col["qty"]] or 0)
        amount = float(row[col["amount"]] or 0)
        source = (row[col["source"]] or "")
        key = (sku, model)
        s = source.lower()
        try:
            txn_date = datetime.strptime(str(row[col["date"]]).strip(), "%d.%m.%Y %H:%M").date()
        except (ValueError, TypeError):
            txn_date = None
        # Отчёт отбирает строки не только по дате операции: в «сентябрьский» попадают
        # и операции августа, которые есть и в «августовском». Для помесячной разбивки
        # это задваивало деньги (сотни операций на стыке месяцев).
        # strict_window оставляет только операции, датированные внутри окна, — тогда
        # каждая операция попадает ровно в один месяц.
        if strict_window and txn_date and not (date_from <= txn_date <= date_to):
            continue
        _acc(out, key, amount, qty, s)
        if out_recent is not None and txn_date and txn_date >= recent_cutoff:
            _acc(out_recent, key, amount, qty, s)

    # Санити: в реальном отчёте о платежах не может быть движения денег при нулевом
    # количестве по ВСЕМ парам — так выглядит именно сбитый разбор (штуки считаются
    # только в ветке «платёж покупателя»). Падаем, чтобы брак не попал в кэш:
    # квартальный живёт 7 дней, а завершённый месяц — вообще навсегда.
    if out and sum(a["qty"] for a in out.values()) == 0 \
            and sum(abs(a["payout"]) for a in out.values()) > 0:
        raise RuntimeError("разбор отчёта «По платежам» сбит: движение денег есть, "
                           "а количество нулевое по всем SKU — в кэш не пишем")
    if unknown_sources:
        print(f"  ⚠️  нераспознанные источники транзакций (ушли в «Прочее»): "
              f"{ {k: round(v) for k, v in unknown_sources.items()} }")

    for a in out.values():
        a["payout"] = round(a["payout"], 2)
        for k in ("buyer_pay", "returns", "services", "promo", "points", "other"):
            a[k] = round(a[k], 2)
    print(f"  Netting ready: {len(out)} (sku,model) pairs")
    if out_recent is not None:
        for a in out_recent.values():
            a["payout"] = round(a["payout"], 2)
            for k in ("buyer_pay", "returns", "services", "promo", "points", "other"):
                a[k] = round(a[k], 2)
        print(f"  Netting recent({recent_days}d): {len(out_recent)} (sku,model) pairs")
        return out, out_recent
    return out
