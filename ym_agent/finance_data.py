"""Status-aware financial records. Cash and non-cash bonus ledgers never mix."""
from .config import DATA_DIR
from collections import defaultdict
from datetime import date, datetime
from pathlib import Path
import io
import json
import time
import openpyxl

SCHEMA_VERSION = 2
CASH_DONE = {'Переведён по графику выплат', 'Удержан из платежей покупателей'}
CASH_SENT = {'Отправлен'}
CASH_PENDING = {'Будет переведён по графику выплат', 'Будет удержан из платежей покупателей'}
BONUS_DONE = {'Справочно: пополнен баланс', 'Справочно: списан с баланса'}
BONUS_PENDING = {'Справочно: будет пополнен баланс', 'Справочно: будет списан с баланса'}

def ledger(status):
    status = ' '.join(str(status or '').split())
    if status.startswith('Не будет ') or status.startswith('Справочно: не будет '):
        return 'excluded'
    if status in CASH_DONE: return 'cash_done'
    if status in CASH_SENT: return 'cash_sent'
    if status in CASH_PENDING: return 'cash_pending'
    if status in BONUS_DONE: return 'bonus_done'
    if status in BONUS_PENDING: return 'bonus_pending'
    return 'unknown'

def as_date(value):
    if isinstance(value, (datetime, date)): return value.isoformat()[:10]
    for fmt in ('%d.%m.%Y %H:%M', '%d.%m.%Y', '%Y-%m-%d'):
        try: return datetime.strptime(str(value).strip(), fmt).date().isoformat()
        except ValueError: pass
    raise ValueError(f'Не распознана дата операции: {value!r}')

def read_netting(content):
    wb = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
    try:
        ws = wb['Отчёт о платежах']
        ws.reset_dimensions()  # Some Yandex files incorrectly declare A1:A1.
        rows = iter(ws.values)
        next(rows)
        headers = [' '.join(str(v or '').split()).lower() for v in next(rows)]
        needles = {'date':'дата транзакции', 'sku':'ваш sku', 'model':'модели работы',
                   'amount':'сумма транзакции', 'status':'статус', 'source':'источник транзакции',
                   'order':'номер заказа или отгрузки', 'service':'название товара', 'qty':'количество'}
        cols = {}
        for key, needle in needles.items():
            hits = [i for i,h in enumerate(headers) if needle in h]
            if len(hits) != 1: raise ValueError(f'Неоднозначная/отсутствующая колонка: {needle}')
            cols[key] = hits[0]
        result = []
        for row in rows:
            if not any(v is not None for v in row): continue
            row = list(row) + [None] * max(0, len(headers)-len(row))
            if row[cols['amount']] is None: continue
            record = {k:row[i] for k,i in cols.items()}
            record['date'] = as_date(record['date'])
            record['amount'] = round(float(record['amount']), 2)
            record['qty'] = float(record['qty']) if record['qty'] is not None else None
            record['order'] = str(int(record['order'])) if isinstance(record['order'], (int,float)) else str(record['order'] or '')
            record['sku'] = str(record['sku'] or '(без SKU)')
            record['model'] = str(record['model'] or '')
            record['ledger'] = ledger(record['status'])
            result.append(record)
        return result
    finally: wb.close()

def summarize(records, date_from, date_to):
    """Signed amounts by transaction date, including charges with no sales."""
    result = {}
    for r in records:
        if not str(date_from) <= r['date'] <= str(date_to): continue
        key = (r['sku'], r['model'])
        a = result.setdefault(key, {k:0.0 for k in
            ('cash_done','cash_sent','cash_pending','bonus_done','bonus_pending','excluded','unknown')})
        a[r['ledger']] += r['amount']
    return {key:{k:round(v,2) for k,v in a.items()} for key,a in result.items()}

def download_report(kind, date_from, date_to):
    from .config import BUSINESS_ID, YM_API_BASE
    from .yandex_client import _headers, _request_raw_with_retry, _wait_for_report
    for attempt in range(8):
        r = _request_raw_with_retry('POST', f'{YM_API_BASE}/reports/{kind}/generate',
            headers=_headers(), json={'businessId':int(BUSINESS_ID),
            'dateFrom':str(date_from),'dateTo':str(date_to)})
        if r.status_code == 420:
            print(f'   {kind}: лимит API, повтор через 130 с', flush=True)
            time.sleep(130)
            continue
        r.raise_for_status()
        url = _wait_for_report(r.json()['result']['reportId'], max_wait_sec=300)
        response = _request_raw_with_retry('GET', url, timeout=180)
        response.raise_for_status()
        return response.content
    raise RuntimeError(f'{kind}: превышен лимит ожидания API')

def get_netting_records(run_date, date_from, date_to):
    directory = Path(DATA_DIR) / 'finance_cache_v2'
    directory.mkdir(exist_ok=True)
    stem = f'netting_{date_from}_{date_to}'
    cache_path = directory / f'{stem}.json'
    if cache_path.exists():
        cached = json.loads(cache_path.read_text())
        if cached.get('schema_version') == SCHEMA_VERSION and cached.get('as_of') == str(run_date):
            return cached['records']
    content = download_report('united-netting', date_from, date_to)
    records = read_netting(content)
    (directory / f'{stem}.xlsx').write_bytes(content)
    temp = cache_path.with_suffix('.tmp')
    temp.write_text(json.dumps({'schema_version':SCHEMA_VERSION,'as_of':str(run_date),
        'from':str(date_from),'to':str(date_to),'records':records},ensure_ascii=False))
    temp.replace(cache_path)
    return records
