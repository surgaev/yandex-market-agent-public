import os

# --- Yandex Market credentials ---
OAUTH_TOKEN = os.environ.get("YM_OAUTH_TOKEN", "")  # Bearer token
BUSINESS_ID = os.environ.get("YM_BUSINESS_ID", "")
CAMPAIGN_ID = os.environ.get("YM_CAMPAIGN_ID", "")       # FBS-кампания
FBY_CAMPAIGN_ID = os.environ.get("YM_FBY_CAMPAIGN_ID", "")  # FBY-кампания

# --- Report output ---
REPORT_FOLDER = os.environ.get("YM_REPORT_FOLDER", "reports")

# --- Alert thresholds ---
# Shows are considered "low" when yesterday < (THRESHOLD_RATIO * rolling_avg_7d).
# 0 shows while PUBLISHED always triggers alert regardless.
THRESHOLD_RATIO = 0.5   # below 50% of 7-day avg → alert

# --- Артикулы, которые не показываем в отчёте ---
# Сняты с продажи. Точное совпадение артикула. Данные по ним по-прежнему
# копятся в кэшах (выплаты, история ставок и остатков), фильтр только на выходе —
# вернуть артикул в отчёт = просто убрать его отсюда.
EXCLUDED_SKUS: set[str] = set()   # например {"SKU-1", "SKU-2"}

# --- API base ---
YM_API_BASE = "https://api.partner.market.yandex.ru"

# --- A/B тест буста (лист «Текущий тест») ---
# Пусто = тест не идёт, лист не строится. Заполни при запуске нового теста: что
# меняли (TEST) и что оставили (CONTROL). Сравнение идёт по БУСТ-показам
# (не суммарным) день-к-дню.
BOOST_TEST_LABEL = ""
BOOST_TEST_START = ""
BOOST_TEST_SKUS = []
BOOST_CONTROL_SKUS = []
