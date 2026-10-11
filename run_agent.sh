#!/bin/bash
# Cron/launchd wrapper for Yandex Market agent
# Schedule: 06:00 UTC = 09:00 MSK

set -e

AGENT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
LOG_FILE="$AGENT_DIR/agent.log"

# ── Load .env ──
if [ -f "$AGENT_DIR/.env" ]; then
  export $(grep -v '^#' "$AGENT_DIR/.env" | xargs)
fi

if [ -z "$YM_OAUTH_TOKEN" ]; then
  echo "[$(date)] ERROR: YM_OAUTH_TOKEN not set" >> "$LOG_FILE"
  exit 1
fi

if [ -z "$YM_REPORT_FOLDER" ]; then
  echo "[$(date)] ERROR: YM_REPORT_FOLDER not set" >> "$LOG_FILE"
  exit 1
fi

# ── Дневной добор (--if-missing): отдельная джоба в 13:00 нужна только если
#    утренний прогон не состоялся (Мак был на батарее, спал или упал). Ручной
#    запуск без флага работает всегда и ничего не проверяет.
if [ "$1" = "--if-missing" ]; then
  if [ -f "$YM_REPORT_FOLDER/ym_report_$(date +%F).xlsx" ]; then
    echo "[$(date)] Отчёт за сегодня уже есть — дневной добор не нужен" >> "$LOG_FILE"
    exit 0
  fi
  echo "[$(date)] Отчёта за сегодня нет — запускаем дневной добор" >> "$LOG_FILE"
fi

# ── Блокировка: один экземпляр на всё время, включая ожидание зарядки ──
# Ожидание питания может длиться часы, и pgrep по main.py ждущую обёртку не
# видит — без блокировки добор в 13:00 или ручной запуск стартовали бы параллельно.
LOCK_DIR="/tmp/ym-agent.lock"
if ! mkdir "$LOCK_DIR" 2>/dev/null; then
  OTHER_PID="$(cat "$LOCK_DIR/pid" 2>/dev/null || true)"
  if [ -n "$OTHER_PID" ] && kill -0 "$OTHER_PID" 2>/dev/null; then
    echo "[$(date)] Другой запуск уже идёт или ждёт зарядку (pid $OTHER_PID) — выходим" >> "$LOG_FILE"
    exit 0
  fi
  echo "[$(date)] Найдена брошенная блокировка (pid ${OTHER_PID:-?} не жив) — забираем" >> "$LOG_FILE"
fi
echo $$ > "$LOCK_DIR/pid"
trap 'rm -rf "$LOCK_DIR"' EXIT

# ── Ожидание питания (на батарее с закрытой крышкой Мак цикличит в Power Nap /
#    Maintenance Sleep, caffeinate -i это не перебивает — сеть держится по
#    несколько секунд, агент 2+ часа бьётся в ретраи и всё равно падает).
#    Поэтому на батарее не стартуем, а ждём зарядку: проверка каждые 15 минут,
#    до POWER_RETRIES раз. Интервал отмеряется по настенным часам — обычный
#    sleep на время сна Мака замирает, и «15 минут» растянулись бы на часы.
#    caffeinate на время ожидания НЕ включаем: держать Мак бодрым на батарее незачем.
POWER_RETRIES=8          # 8 × 15 мин = до 2 часов ожидания
POWER_INTERVAL=900
attempt=0
while ! pmset -g batt | head -1 | grep -q "AC Power"; do
  if [ $attempt -ge $POWER_RETRIES ]; then
    echo "[$(date)] ERROR: Мак так и не подключили к зарядке за $((POWER_RETRIES * POWER_INTERVAL / 60)) мин — прогон отменён" >> "$LOG_FILE"
    /usr/bin/osascript -e 'display notification "Мак так и не подключили к зарядке — прогон отчёта отменён, подключи питание и запусти вручную" with title "YM Agent: на батарее" sound name "Basso"' 2>/dev/null || true
    exit 1
  fi
  attempt=$((attempt + 1))
  echo "[$(date)] Мак на батарее — повторная проверка через $((POWER_INTERVAL / 60)) мин ($attempt/$POWER_RETRIES)" >> "$LOG_FILE"
  next_check=$(( $(date +%s) + POWER_INTERVAL ))
  while [ "$(date +%s)" -lt "$next_check" ]; do
    sleep 30
  done
done
if [ $attempt -gt 0 ]; then
  echo "[$(date)] Питание подключено после $attempt ожиданий — стартуем" >> "$LOG_FILE"
fi

# ── Ожидание сети (Mac мог только что проснуться) ──
# Проверяем curl'ом по HTTPS, а не пингом: VPN и фаерволы вполне пропускают ICMP,
# но режут TCP — тогда пинг проходит, проверка считает сеть живой, а прогон всё
# равно умирает через два часа ретраев. Любой HTTP-ответ (в т.ч. 401 без токена)
# = путь до API рабочий.
API_URL="https://api.partner.market.yandex.ru"
DIAG_INTERFACE="en0"   # только для диагностики; трафик через него НЕ гоним
MAX_WAIT=120
WAITED=0
while ! curl -s -m 5 -o /dev/null "$API_URL" 2>/dev/null; do
  if [ $WAITED -ge $MAX_WAIT ]; then
    # Кто виноват: нет сети вообще или VPN увёл маршрут в туннель? В конце
    # августа три прогона подряд умерли именно из-за VPN, режущего .ru, и на
    # выяснение этого ушло несколько дней — теперь причина пишется сразу.
    DEF_IFACE="$(route -n get default 2>/dev/null | awk '/interface:/{print $2}')"
    if curl -s -m 5 --interface "$DIAG_INTERFACE" -o /dev/null "$API_URL" 2>/dev/null; then
      echo "[$(date)] ERROR: API недоступен через маршрут по умолчанию (${DEF_IFACE:-?}), но отвечает через ${DIAG_INTERFACE} — похоже, мешает VPN. Отключи VPN и запусти вручную: bash run_agent.sh" >> "$LOG_FILE"
      /usr/bin/osascript -e 'display notification "API Яндекса не отвечает через VPN, но доступен напрямую — отключи VPN и запусти вручную" with title "YM Agent: мешает VPN" sound name "Basso"' 2>/dev/null || true
    else
      echo "[$(date)] ERROR: Нет сети после ${MAX_WAIT}s (маршрут по умолчанию: ${DEF_IFACE:-нет}), прогон отменён" >> "$LOG_FILE"
      /usr/bin/osascript -e 'display notification "Нет сети — прогон отчёта отменён" with title "YM Agent: нет сети" sound name "Basso"' 2>/dev/null || true
    fi
    exit 1
  fi
  sleep 5
  WAITED=$((WAITED + 5))
done

if [ $WAITED -gt 0 ]; then
  echo "[$(date)] Network came up after ${WAITED}s" >> "$LOG_FILE"
fi

echo "[$(date)] Starting Yandex Market agent..." >> "$LOG_FILE"

cd "$AGENT_DIR"

if [ -d "$AGENT_DIR/venv" ]; then
  source "$AGENT_DIR/venv/bin/activate"
fi

# Не даём set -e молча убить скрипт: при падении python — уведомление на Mac,
# чтобы сбой был виден сразу, а не когда хватишься отчёта.
if caffeinate -i python3 main.py >> "$LOG_FILE" 2>&1; then
  echo "[$(date)] Agent finished." >> "$LOG_FILE"
else
  echo "[$(date)] Agent FAILED (см. лог выше)" >> "$LOG_FILE"
  /usr/bin/osascript -e 'display notification "Отчёт Яндекс.Маркета НЕ собрался — смотри agent.log" with title "YM Agent: сбой" sound name "Basso"' 2>/dev/null || true
  exit 1
fi
