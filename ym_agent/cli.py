"""Команда `ym-agent`: init — создать настройки, run — собрать отчёт."""
import argparse
import os
import sys
from datetime import date

DATA_DIR = os.environ.get("YM_AGENT_HOME", os.path.expanduser("~/.ym-agent"))
ENV_FILE = os.path.join(DATA_DIR, ".env")
COST_FILE = os.path.join(DATA_DIR, "cost_prices.csv")

ENV_TEMPLATE = """\
# Ключ API: кабинет продавца → Настройки → API и модули → Токены авторизации.
# Доступы: статистика, финансы, заказы, товары/остатки, продвижение.
YM_OAUTH_TOKEN=
# ID бизнеса и кампаний — в адресной строке кабинета (business=…, campaign=…).
YM_BUSINESS_ID=
YM_CAMPAIGN_ID=        # FBS — ваш склад
YM_FBY_CAMPAIGN_ID=    # FBY — склад Яндекса
# Куда класть отчёты (по умолчанию ~/Downloads):
# YM_REPORT_FOLDER=/Users/you/Downloads
"""
COST_TEMPLATE = """\
# Закупочные цены, ₽. Артикул или его префикс: ABC123 покрывает ABC123-red, ABC123-xl.
# Разделитель — запятая или точка с запятой; можно редактировать в Excel.
sku,cost
ABC123,100
"""


def _load_env():
    if not os.path.exists(ENV_FILE):
        return
    for line in open(ENV_FILE, encoding="utf-8"):
        line = line.split("#", 1)[0].strip()
        if "=" in line:
            k, v = line.split("=", 1)
            if v.strip():
                os.environ.setdefault(k.strip(), v.strip())


def _bind_en0():
    """Исходящие соединения — через Wi-Fi (en0) в обход VPN, который режет API (macOS)."""
    import socket
    import subprocess
    import urllib3.connection as u3c
    import urllib3.util.connection as u3
    ip = subprocess.check_output(["ipconfig", "getifaddr", "en0"]).decode().strip()
    orig = u3.create_connection

    def create_connection(address, timeout=socket._GLOBAL_DEFAULT_TIMEOUT, source_address=None, **kw):
        return orig(address, timeout, (ip, 0), **kw)
    u3.create_connection = create_connection
    u3c.connection.create_connection = create_connection
    print(f"исходящие соединения привязаны к en0 ({ip})")


def cmd_init(_args):
    os.makedirs(DATA_DIR, exist_ok=True)
    for path, body in ((ENV_FILE, ENV_TEMPLATE), (COST_FILE, COST_TEMPLATE)):
        if os.path.exists(path):
            print(f"  уже есть: {path}")
        else:
            open(path, "w", encoding="utf-8").write(body)
            print(f"  создан:   {path}")
    print(f"\nДальше:\n  1. впишите ключ и ID кабинета в {ENV_FILE}"
          f"\n  2. впишите закупочные цены в {COST_FILE}"
          f"\n  3. запустите: ym-agent run")


def cmd_run(args):
    _load_env()
    missing = [k for k in ("YM_OAUTH_TOKEN", "YM_BUSINESS_ID", "YM_CAMPAIGN_ID") if not os.environ.get(k)]
    if missing:
        sys.exit(f"Не заполнено в {ENV_FILE}: {', '.join(missing)}. Сначала: ym-agent init")
    if not os.path.exists(COST_FILE):
        print(f"⚠️  нет {COST_FILE} — себестоимость будет 0, прибыль завышена")
    if args.en0:
        _bind_en0()
    from . import main   # импорт после .env: config читает переменные при загрузке
    main.run(date.fromisoformat(args.date) if args.date else date.today())


def main():
    p = argparse.ArgumentParser(prog="ym-agent", description="Аналитика продавца Яндекс Маркета → Excel-отчёт")
    sub = p.add_subparsers(dest="cmd", required=True)
    sub.add_parser("init", help=f"создать шаблоны настроек в {DATA_DIR}").set_defaults(func=cmd_init)
    r = sub.add_parser("run", help="собрать отчёт")
    r.add_argument("--date", help="дата отчёта YYYY-MM-DD (по умолчанию сегодня)")
    r.add_argument("--en0", action="store_true", help="macOS: в обход VPN через Wi-Fi (en0)")
    r.set_defaults(func=cmd_run)
    args = p.parse_args()
    args.func(args)


if __name__ == "__main__":
    main()
