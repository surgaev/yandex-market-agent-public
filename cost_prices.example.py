"""
Cost prices (закуп) per SKU prefix.

Source: Google Sheets "<ID вашей таблицы>",
(синхронизировано с таблицей 20.09.2026; вручную заданные префиксы,
которых в таблице нет, оставлены как есть — см. комментарии ниже)
sheet "Sheet2", column I (закуп).
Variants share cost via prefix matching (SKUzbl → SKU).
"""

_COST_MAP = {
    # префикс артикула: закупочная цена, ₽
    "SKU": 100.0,   # пример: все SKU* берут эту цену по префиксу
}

TAX_RATE = 0.07


def get_cost_price(sku: str) -> float | None:
    if sku in _COST_MAP:
        return _COST_MAP[sku]
    best_match = None
    best_len = 0
    for prefix, cost in _COST_MAP.items():
        if sku.startswith(prefix) and len(prefix) > best_len:
            best_match = cost
            best_len = len(prefix)
    return best_match
