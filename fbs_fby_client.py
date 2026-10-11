"""
FBS / FBY data: orders and stocks.

Campaigns:
  FBS = YM_CAMPAIGN_ID, FBY = YM_FBY_CAMPAIGN_ID (из .env)

Stock types:
  FIT        — годный к продаже (на складе)
  AVAILABLE  — доступный покупателю (FIT − FREEZE)
  FREEZE     — зарезервированный под заказ
  QUARANTINE — на карантине (проверка/брак)
"""

import time
from datetime import date, timedelta
from collections import defaultdict
from config import OAUTH_TOKEN, BUSINESS_ID, YM_API_BASE, CAMPAIGN_ID, FBY_CAMPAIGN_ID
from yandex_client import _request_raw_with_retry

FBS_CAMPAIGN = CAMPAIGN_ID
FBY_CAMPAIGN = FBY_CAMPAIGN_ID


def _headers():
    return {"Api-Key": OAUTH_TOKEN, "Content-Type": "application/json"}


def _post(path, body, params=None, timeout=60, retries=5):
    """POST с ретраем сетевых сбоев (DNS/connection/timeout) — см. yandex_client."""
    r = _request_raw_with_retry("POST", f"{YM_API_BASE}{path}", headers=_headers(),
                                 json=body, params=params, timeout=timeout,
                                 retries=retries)
    r.raise_for_status()
    return r.json()


# ── Orders ────────────────────────────────────────────────────────────────────

def _fetch_orders(campaign_id: str, date_from: date, date_to: date) -> list[dict]:
    """Fetch orders via POST /campaigns/{id}/stats/orders. Returns raw order list."""
    all_orders = []
    page_token = None
    seen = set()
    while True:
        params = {"pageToken": page_token} if page_token else {}
        data = _post(f"/campaigns/{campaign_id}/stats/orders",
                      {"dateFrom": date_from.isoformat(), "dateTo": date_to.isoformat()},
                      params=params)["result"]
        all_orders.extend(data.get("orders", []))
        page_token = data.get("paging", {}).get("nextPageToken")
        if not page_token or page_token in seen:
            break
        seen.add(page_token)
        time.sleep(0.2)
    return all_orders


def _orders_to_daily(orders: list[dict]) -> dict[str, dict[str, int]]:
    """
    Group orders by date → shopSku → total item count.
    Returns: { "YYYY-MM-DD": { "shopSku": count } }
    """
    result = defaultdict(lambda: defaultdict(int))
    for o in orders:
        dt = o.get("creationDate", "")
        for item in o.get("items", []):
            sku = item.get("shopSku", "")
            cnt = item.get("count", 1)
            result[dt][sku] += cnt
    return dict(result)


def get_orders_by_type(date_from: date, date_to: date) -> dict:
    """
    Returns:
    {
      "fbs": { "YYYY-MM-DD": { "shopSku": count } },
      "fby": { "YYYY-MM-DD": { "shopSku": count } },
    }
    """
    fbs_orders = _fetch_orders(FBS_CAMPAIGN, date_from, date_to)
    fby_orders = _fetch_orders(FBY_CAMPAIGN, date_from, date_to)
    return {
        "fbs": _orders_to_daily(fbs_orders),
        "fby": _orders_to_daily(fby_orders),
    }


# ── Orders with commissions (unit economics) ─────────────────────────────────

COMMISSION_TYPES = [
    "FEE",                    # размещение на витрине
    "AUCTION_PROMOTION",      # буст продаж
    "DELIVERY_TO_CUSTOMER",   # доставка покупателю
    "CROSSREGIONAL_DELIVERY", # средняя миля
    "PAYMENT_TRANSFER",       # приём + перевод платежа
    "AGENCY",                 # агентское вознаграждение
    "SORTING",                # обработка заказов в СЦ или ПВЗ
    "FULFILLMENT",            # фулфилмент
    "RETURN_PROCESSING",      # обработка возвратов
]

DELIVERED_STATUSES = {"DELIVERED"}
RETURN_STATUSES = {
    "RETURNED", "CANCELLED_IN_DELIVERY", "CANCELLED_IN_PROCESSING",
    "CANCELLED_BEFORE_PROCESSING",
}


def get_orders_with_fees(date_from: date, date_to: date) -> dict:
    """
    Fetch orders with per-item commission breakdown from both FBS and FBY.
    Filters by statusUpdateDate (delivery/return date), not creationDate.
    Fetches a wider window (45 days) to catch orders created earlier but
    delivered/returned within the target period.

    Commissions are split proportionally across items by buyer_price*count.

    Returns: {
        "fbs": [item_dict, ...],
        "fby": [item_dict, ...],
    }
    Each item_dict has per-UNIT commission values (order comms / item_count).
    """
    fetch_from = date_from - timedelta(days=30)
    result = {"fbs": [], "fby": []}
    for label, campaign_id in [("fbs", FBS_CAMPAIGN), ("fby", FBY_CAMPAIGN)]:
        raw = _fetch_orders(campaign_id, fetch_from, date_to)
        for o in raw:
            status = o.get("status", "")
            if status not in DELIVERED_STATUSES and status not in RETURN_STATUSES:
                continue

            # Filter by statusUpdateDate within target period
            sud = o.get("statusUpdateDate", "")
            if sud:
                sud_date = date.fromisoformat(sud[:10])
                if sud_date < date_from or sud_date > date_to:
                    continue

            comm_map = {}
            for c in o.get("commissions", []):
                comm_map[c["type"]] = c.get("actual", 0)
            order_total_fees = sum(comm_map.values())

            items = o.get("items", [])
            # Total order revenue for proportional splitting
            order_revenue = 0
            item_revenues = []
            for item in items:
                bp = 0
                for p in item.get("prices", []):
                    if p["type"] == "BUYER":
                        bp = p.get("total", 0)  # total = costPerItem * count
                item_revenues.append(bp)
                order_revenue += bp

            for idx, item in enumerate(items):
                buyer_price_per_unit = 0
                marketplace_price_per_unit = 0
                for p in item.get("prices", []):
                    if p["type"] == "BUYER":
                        buyer_price_per_unit = p.get("costPerItem", 0)
                    elif p["type"] == "MARKETPLACE":
                        marketplace_price_per_unit = p.get("costPerItem", 0)

                cnt = item.get("count", 1)
                # Proportional share of order-level commissions for this item
                share = item_revenues[idx] / order_revenue if order_revenue > 0 else 1.0 / len(items)
                item_comms = {k: round(v * share / cnt, 2) for k, v in comm_map.items()}
                item_total_fees = round(order_total_fees * share / cnt, 2)

                region = o.get("deliveryRegion", {})
                result[label].append({
                    "order_id": o.get("id"),
                    "date": o.get("creationDate", ""),
                    "status_update_date": sud[:10] if sud else "",
                    "status": status,
                    "sku": item.get("shopSku", ""),
                    "name": item.get("offerName", ""),
                    "delivery_region": region.get("name", ""),
                    "delivery_region_id": region.get("id", 0),
                    "count": cnt,
                    "buyer_price": buyer_price_per_unit,
                    "marketplace_price": marketplace_price_per_unit,
                    "commissions": item_comms,
                    "total_fees": item_total_fees,
                })
    return result


def get_promo_prices() -> dict[str, dict]:
    """
    Fetch promo (discount) prices for offers participating in active promos.

    Returns: { "shopSku": {"promo_price": float, "regular_price": float, "promo_name": str,
                           "promo_id": str, "period_from": str, "period_to": str,
                           "all_promos": [{"promo_id","promo_name","promo_price",
                                          "regular_price","period_from","period_to"}, ...]} }
    Only includes offers from promos where participating == true.

    Топ-уровневые поля (promo_price/regular_price/...) — ЭФФЕКТИВНАЯ акция (самая
    дешёвая), она же реальная цена на витрине для покупателя: Яндекс показывает
    минимум, если артикул в нескольких акциях. Для "во что артикул реально
    участвует" (история, лист «Ставки буста по дням») используй `all_promos` —
    один SKU часто в 2+ акциях ОДНОВРЕМЕННО (проверено на живых данных: 14 из 26
    активных SKU сразу в двух), раньше это схлопывалось в одну и терялось.

    promo_id/period_* нужны для истории участия по дням (promo_daily.json): API
    отдаёт только ТЕКУЩЕЕ участие, истории нет — накапливаем сами, как ставки буста.
    """
    result = {}
    all_by_sku: dict[str, list] = {}
    failed: list[str] = []

    promos = _post(f"/businesses/{BUSINESS_ID}/promos", {})["result"].get("promos", [])

    for promo in promos:
        if not promo.get("participating"):
            continue
        promo_id = promo["id"]
        promo_name = promo.get("name", "")
        period = promo.get("period") or {}
        p_from = str(period.get("dateTimeFrom", ""))[:10]
        p_to = str(period.get("dateTimeTo", ""))[:10]

        try:
            _collect_promo_offers(promo_id, promo_name, p_from, p_to, result, all_by_sku)
        except Exception as e:
            # 17.09.2026: у только что стартовавшей «Бестселлеры: Новая единая акция»
            # /promos/offers стабильно отдавал 500 — и одна сломанная акция роняла
            # весь прогон. Пропускаем её, остальные акции собираем как обычно.
            print(f"   ⚠️  офферы акции {promo_id} «{promo_name}» не загрузились: {str(e)[:120]}")
            failed.append(f"{promo_id} «{promo_name}»")

    for sku, entry in result.items():
        # без copy() правки в all_promos ниже не тронут исходный entry
        result[sku] = dict(entry)
        result[sku]["all_promos"] = sorted(all_by_sku.get(sku, [entry]),
                                           key=lambda e: e["promo_id"])

    if failed:
        raise PromoFetchPartial(result, failed)
    return result


class PromoFetchPartial(Exception):
    """Часть акций не загрузилась. result — что удалось собрать, failed — какие акции выпали.
    Отдельное исключение, а не тихий неполный результат: вызывающему нужно знать,
    что «артикула нет в акции» может означать «акцию не удалось прочитать»."""
    def __init__(self, result: dict, failed: list[str]):
        super().__init__("не загрузились акции: " + ", ".join(failed))
        self.result = result
        self.failed = failed


def _collect_promo_offers(promo_id, promo_name, p_from, p_to, result, all_by_sku):
    """Все страницы офферов одной акции → дописывает в result / all_by_sku на месте."""
    page_token = None
    seen = set()
    while True:
        params = {"pageToken": page_token} if page_token else {}
        data = _post(f"/businesses/{BUSINESS_ID}/promos/offers",
                     {"promoId": promo_id}, params=params)["result"]
        for off in data.get("offers", []):
            sku = off.get("offerId", "")
            dp = off.get("params", {}).get("discountParams", {})
            promo_price = dp.get("promoPrice")
            regular_price = dp.get("price")
            if promo_price is None:
                continue
            entry = {
                "promo_price": promo_price,
                "regular_price": regular_price,
                "promo_name": promo_name,
                "promo_id": promo_id,
                "period_from": p_from,
                "period_to": p_to,
            }
            all_by_sku.setdefault(sku, []).append(entry)
            existing = result.get(sku)
            if existing is None or promo_price < existing["promo_price"]:
                result[sku] = entry
        page_token = data.get("paging", {}).get("nextPageToken")
        if not page_token or page_token in seen:
            break
        seen.add(page_token)
        time.sleep(0.2)


def get_catalog_prices() -> dict[str, float]:
    """Fetch basicPrice for all offers from catalog. Returns {shopSku: price}."""
    prices = {}
    page_token = None
    seen = set()
    while True:
        params = {"pageToken": page_token} if page_token else {}
        data = _post(f"/businesses/{BUSINESS_ID}/offer-mappings", {}, params=params)["result"]
        for item in data.get("offerMappings", []):
            offer = item.get("offer", {})
            sku = offer.get("offerId", "")
            bp = offer.get("basicPrice", {})
            if bp.get("value"):
                prices[sku] = bp["value"]
        page_token = data.get("paging", {}).get("nextPageToken")
        if not page_token or page_token in seen:
            break
        seen.add(page_token)
        time.sleep(0.2)
    return prices


# ── Price competitiveness (индекс/привлекательность цен) ───────────────────────

def get_price_recommendations() -> dict[str, dict]:
    """
    Привлекательность цен по каждому SKU (раздел «Цены» в кабинете).
    POST /businesses/{id}/offers/recommendations.

    Returns: { shopSku: {price, competitiveness, shows, optimal_price, average_price, band} }
      competitiveness — enum Маркета: OPTIMAL / AVERAGE / NONCOMPETITIVE;
                        у low-traffic офферов отсутствует (None)
      optimal_price   — рекомендованная цена, чтобы быть/остаться «выгодным»
      average_price   — средняя цена рынка по модели (ориентир)
      band            — метка из enum: «Выгодная»/«Умеренная»/«Завышенная»/«—»
    """
    BAND = {"OPTIMAL": "Выгодная", "AVERAGE": "Умеренная",
            "NONCOMPETITIVE": "Завышенная", "NON_COMPETITIVE": "Завышенная"}
    out = {}
    page_token = None
    seen = set()
    while True:
        params = {"pageToken": page_token} if page_token else {}
        data = _post(f"/businesses/{BUSINESS_ID}/offers/recommendations", {}, params=params)["result"]
        for item in data.get("offerRecommendations", []):
            offer = item.get("offer", {})
            sku = offer.get("offerId", "")
            if not sku:
                continue
            price = (offer.get("price") or {}).get("value")
            th = (item.get("recommendation") or {}).get("competitivenessThresholds", {})
            optimal = (th.get("optimalPrice") or {}).get("value")
            average = (th.get("averagePrice") or {}).get("value")
            comp = offer.get("competitiveness")
            band = BAND.get(comp, "—")
            out[sku] = {
                "price": price,
                "competitiveness": comp,
                "shows": offer.get("shows", 0),
                "optimal_price": optimal,
                "average_price": average,
                "band": band,
            }
        page_token = data.get("paging", {}).get("nextPageToken")
        if not page_token or page_token in seen:
            break
        seen.add(page_token)
        time.sleep(0.2)
    return out


# ── Reviews (отзывы и рейтинг) ─────────────────────────────────────────────────

def get_reviews_summary() -> dict[str, dict]:
    """
    Агрегированные отзывы по SKU. POST /businesses/{id}/goods-feedback (пагинация).

    Returns: { shopSku: {count, avg_rating, unanswered, r5,r4,r3,r2,r1, last_date,
                         last_neg} }
      unanswered — отзывы с needReaction=true (продавец не ответил)
      last_neg   — текст последнего отзыва с рейтингом ≤3 (или "")
    """
    agg = defaultdict(lambda: {"count": 0, "sum": 0, "unanswered": 0,
                               "r5": 0, "r4": 0, "r3": 0, "r2": 0, "r1": 0,
                               "last_date": "", "last_neg": ""})
    page_token = None
    seen = set()
    while True:
        params = {"pageToken": page_token} if page_token else {}
        # best-effort фаза: при сбое сети скипаем быстро (retries=2)
        data = _post(f"/businesses/{BUSINESS_ID}/goods-feedback", {}, params=params,
                     retries=2)["result"]
        for fb in data.get("feedbacks", []):
            sku = (fb.get("identifiers") or {}).get("offerId", "")
            if not sku:
                continue
            a = agg[sku]
            rating = (fb.get("statistics") or {}).get("rating")
            a["count"] += 1
            if rating:
                a["sum"] += rating
                a[f"r{int(rating)}"] = a.get(f"r{int(rating)}", 0) + 1
            if fb.get("needReaction"):
                a["unanswered"] += 1
            created = (fb.get("createdAt") or "")[:10]
            if created > a["last_date"]:
                a["last_date"] = created
            if rating and rating <= 3 and not a["last_neg"]:
                desc = fb.get("description") or {}
                a["last_neg"] = (desc.get("comment") or desc.get("disadvantages") or "")[:120]
        page_token = data.get("paging", {}).get("nextPageToken")
        if not page_token or page_token in seen:
            break
        seen.add(page_token)
        time.sleep(0.2)

    out = {}
    for sku, a in agg.items():
        out[sku] = {
            "count": a["count"],
            "avg_rating": round(a["sum"] / a["count"], 2) if a["count"] else 0,
            "unanswered": a["unanswered"],
            "r5": a["r5"], "r4": a["r4"], "r3": a["r3"], "r2": a["r2"], "r1": a["r1"],
            "last_date": a["last_date"],
            "last_neg": a["last_neg"],
        }
    return out


# ── Stocks ────────────────────────────────────────────────────────────────────

def _fetch_stocks(campaign_id: str) -> dict[str, dict[str, int]]:
    """
    Returns: { "shopSku": { "fit": N, "available": N, "freeze": N, "quarantine": N } }
    Aggregated across all warehouses.
    """
    stocks = defaultdict(lambda: {"fit": 0, "available": 0, "freeze": 0, "quarantine": 0})
    page_token = None
    seen = set()
    while True:
        params = {"pageToken": page_token} if page_token else {}
        data = _post(f"/campaigns/{campaign_id}/offers/stocks", {}, params=params)["result"]
        for wh in data.get("warehouses", []):
            for off in wh.get("offers", []):
                sku = off["offerId"]
                for s in off.get("stocks", []):
                    stype = s["type"].lower()
                    if stype in stocks[sku]:
                        stocks[sku][stype] += s["count"]
        page_token = data.get("paging", {}).get("nextPageToken")
        if not page_token or page_token in seen:
            break
        seen.add(page_token)
        time.sleep(0.2)
    return dict(stocks)


def get_stocks() -> dict:
    """
    Returns:
    {
      "fbs": { "shopSku": {fit, available, freeze, quarantine} },
      "fby": { "shopSku": {fit, available, freeze, quarantine} },
    }
    """
    return {
        "fbs": _fetch_stocks(FBS_CAMPAIGN),
        "fby": _fetch_stocks(FBY_CAMPAIGN),
    }


def estimate_fby_restock(
    fby_stocks: dict[str, dict],
    fby_orders_7d: dict[str, dict],
    num_days: int = 7,
) -> list[dict]:
    """
    Estimate days until FBY stock runs out and recommend restock.

    Returns list of {
        sku, available, avg_daily_sales, days_left, restock_qty, urgency
    }
    sorted by days_left ascending.
    """
    # Aggregate 7-day sales per SKU
    sales_total = defaultdict(int)
    for day_data in fby_orders_7d.values():
        for sku, cnt in day_data.items():
            sales_total[sku] += cnt

    results = []
    all_skus = set(fby_stocks.keys()) | set(sales_total.keys())

    for sku in all_skus:
        st = fby_stocks.get(sku, {})
        available = st.get("available", 0)
        total_sold = sales_total.get(sku, 0)
        avg_daily = total_sold / num_days if num_days else 0

        if avg_daily > 0:
            days_left = available / avg_daily
        elif available > 0:
            days_left = 999  # has stock, no sales
        else:
            days_left = 0

        # Restock for 30 days
        restock_qty = max(0, round(avg_daily * 30 - available))

        if days_left <= 3:
            urgency = "🔴 СРОЧНО"
        elif days_left <= 7:
            urgency = "🟡 Скоро"
        elif days_left <= 14:
            urgency = "🔵 Планировать"
        else:
            urgency = "🟢 OK"

        results.append({
            "sku": sku,
            "available": available,
            "avg_daily": round(avg_daily, 2),
            "days_left": round(days_left, 1),
            "restock_qty": restock_qty,
            "urgency": urgency,
        })

    results.sort(key=lambda x: x["days_left"])
    return results
