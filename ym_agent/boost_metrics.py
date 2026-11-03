"""Boost percentages come from boost fields, never promotion participation."""
import math

def number(value):
    try:
        result = float(value)
        return result if math.isfinite(result) else None
    except (TypeError, ValueError): return None

def percentage(value):
    result = number(value)
    return result if result is not None and 0 <= result <= 100 else None

def estimate(total, pct):
    total, pct = number(total), percentage(pct)
    if total is None or total < 0: return None
    if total == 0: return 0.0
    return total * pct / 100 if pct is not None else None

def weighted_share(rows, count_key, percent_key):
    total = boosted = 0.0
    for row in rows:
        count = number(row.get(count_key))
        # Missing counts are not proof of zero activity.
        if count is None: return None
        part = estimate(count, row.get(percent_key))
        if part is None: return None
        total += count
        boosted += part
    return round(100 * boosted / total, 2) if total else None
