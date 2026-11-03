"""Order-linked cash contribution, distinct from nominal price and bonus credit."""
from .config import DATA_DIR
from pathlib import Path
from collections import defaultdict
import io
import json
import openpyxl
from .finance_data import as_date, download_report, SCHEMA_VERSION
from .cost_prices import get_cost_price, TAX_RATE

def _rows(wb, name, required):
    ws=wb[name]; ws.reset_dimensions()
    iterator=iter(ws.values)
    for row in iterator:
        header=[' '.join(str(v or '').split()) for v in row]
        if all(h in header for h in required):
            for h in required:
                if header.count(h)!=1: raise ValueError(f'Неоднозначная колонка {h}')
            return header, list(iterator)
    raise ValueError(f'Не найдена шапка листа {name}')

def read_orders(content):
    wb=openpyxl.load_workbook(io.BytesIO(content),read_only=True,data_only=True)
    try:
        sh,sr=_rows(wb,'Услуги и маржа по заказам',
            ['Номер заказа','Все услуги Маркета за заказы, ₽','Статус заказа'])
        ih,ir=_rows(wb,'Транзакции по заказам и товарам',
            ['Номер заказа','Ваш SKU','Доставлено или возвращено','Статус товара'])
        def get(row,headers,label):
            i=headers.index(label)
            return row[i] if i<len(row) else None
        def key(row,header):
            return (str(get(row,header,'Модели работы')),
                    str(int(get(row,header,'Номер заказа'))))
        services={}
        for r in sr:
            if get(r,sh,'Номер заказа') is None: continue
            k=key(r,sh)
            if k in services: raise ValueError('Повторная строка услуг заказа '+k[1])
            services[k]={'fees':get(r,sh,'Все услуги Маркета за заказы, ₽'),
                         'status':get(r,sh,'Статус заказа')}
        groups=defaultdict(list)
        for r in ir:
            if get(r,ih,'Номер заказа') is None: continue
            k=key(r,ih)
            groups[k].append(r)
        output=[]
        for k,items in groups.items():
            seen=set()
            weights=[]
            for r in items:
                sku=get(r,ih,'Ваш SKU')
                if sku in seen: raise ValueError('Повторный SKU в заказе '+k[1])
                seen.add(sku)
                price=get(r,ih,'Цена продажи (за шт.), ₽')
                count=get(r,ih,'Передано в доставку')
                weights.append(float(price or 0)*float(count or 0))
            fee_total=services.get(k,{}).get('fees')
            allocated=0.0
            for i,r in enumerate(items):
                g=lambda name:get(r,ih,name)
                sku=str(g('Ваш SKU'))
                status=str(g('Статус товара'))
                delivered=status=='Доставлен покупателю'
                cancelled=status=='Отменён' or status.startswith('Невыкуп ')
                qty=float(g('Доставлено или возвращено')) if delivered and g('Доставлено или возвращено') is not None else (0.0 if cancelled else None)
                fee=None
                if fee_total is not None:
                    if len(items)==1: fee=float(fee_total)
                    elif sum(weights)>0:
                        fee=round(float(fee_total)-allocated,2) if i==len(items)-1 else round(float(fee_total)*weights[i]/sum(weights),2)
                        allocated+=fee
                # The first payment group is buyer cash; the next three groups
                # are separate monetary compensation, not bonus credit.
                payment_cols=[j for j,h in enumerate(ih) if h=='Сумма платежа']
                if len(payment_cols)!=4: raise ValueError('Изменилась структура денежных платежей')
                values=[r[j] if j<len(r) else None for j in payment_cols]
                buyer=float(values[0]) if delivered and values[0] is not None else (0.0 if cancelled else None)
                other=sum(float(v or 0) for v in values[1:]) if delivered else (0.0 if cancelled else None)
                bonus=float(g('Сумма баллов') or 0) if g('Статус (справочно)')=='Пополнен баланс' else 0.0
                cost_u=get_cost_price(sku)
                cost=round(cost_u*qty,2) if cost_u is not None and qty is not None else (0.0 if qty==0 else None)
                reason='Рассчитано' if delivered or cancelled else 'Возврат / незавершённый заказ: нужна сверка'
                if cost is None: reason='Неизвестна себестоимость или итоговое количество'
                if fee is None: reason='Нет услуг или базы распределения по SKU'
                margin=reserve=after=None
                # A missing buyer payment is not a free item.
                if buyer is not None and other is not None and fee is not None and cost is not None:
                    margin=round(buyer+other-fee-cost,2)
                    reserve=round((buyer+other)*TAX_RATE,2)
                    after=round(margin-reserve,2)
                elif delivered and buyer is None: reason='Нет суммы платежа покупателя'
                output.append(dict(sku=sku,order=k[1],model=k[0],created=as_date(g('Дата оформления')),
                    status=status,qty=qty,buyer=buyer,other=other,bonus=bonus,fees=fee,cost=cost,
                    contribution=margin,tax_reserve=reserve,after_reserve=after,quality=reason,
                    allocation='по цене × количеству' if len(items)>1 else 'прямо по заказу'))
        return output
    finally: wb.close()

def get_order_records(run_date,date_from,date_to):
    directory=Path(DATA_DIR)/'finance_cache_v2';directory.mkdir(exist_ok=True)
    path=directory/f'orders_{date_from}_{date_to}.json'
    if path.exists():
        cached=json.loads(path.read_text())
        if cached.get('schema_version')==SCHEMA_VERSION and cached.get('as_of')==str(run_date): return cached['records']
    content=download_report('united-orders',date_from,date_to)
    records=read_orders(content)
    path.with_suffix('.xlsx').write_bytes(content)
    temp=path.with_suffix('.tmp')
    temp.write_text(json.dumps({'schema_version':SCHEMA_VERSION,'as_of':str(run_date),
        'from':str(date_from),'to':str(date_to),'records':records},ensure_ascii=False))
    temp.replace(path)
    return records

def style_order_economics(ws):
    """Colour the existing order view without changing its values or layout."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    ws.sheet_properties.tabColor = '287D6B'
    ws.sheet_view.showGridLines = False
    groups = {
        1:('23465A','F3F6F8'),2:('23465A','F3F6F8'),3:('23465A','F3F6F8'),
        4:('23465A','F3F6F8'),5:('23465A','F3F6F8'),6:('23465A','F3F6F8'),
        7:('287D6B','E8F4EF'),8:('287D6B','E8F4EF'),9:('7954A0','F1EAF7'),
        10:('A96325','FFF0E0'),11:('A96325','FFF0E0'),12:('365C85','EAF0F8'),
        13:('A96325','FFF0E0'),14:('365C85','EAF0F8'),
        15:('596773','F3F6F8'),16:('596773','F3F6F8')}
    for c in ws[2]:
        c.fill=PatternFill('solid',fgColor=groups[c.column][0])
        c.font=Font(name='Arial',size=10,bold=True,color='FFFFFF')
        c.alignment=Alignment(wrap_text=True,horizontal='center',vertical='center')
    ws.row_dimensions[2].height=52
    ws['A1'].font=Font(name='Arial',size=10,color='555555')
    line=Side(style='thin',color='DFE5E9')
    for row in ws.iter_rows(min_row=3):
        for c in row:
            c.fill=PatternFill('solid',fgColor=groups[c.column][1])
            c.font=Font(name='Arial',size=10,color='333333',bold=c.column==14)
            c.border=Border(bottom=line)
            c.alignment=Alignment(vertical='center',horizontal='right' if 6<=c.column<=14 else 'left',wrap_text=c.column in (5,15,16))
            if 7<=c.column<=14 and isinstance(c.value,(int,float)):
                c.number_format='#,##0.00;[Red]-#,##0.00;"—"'
        row[5].number_format='#,##0'
        for column in (12,14):
            c=ws.cell(row[0].row,column)
            if isinstance(c.value,(int,float)):
                fill,color=('DFF0E4','1B5E20') if c.value>0 else ('FBE3E2','A61C1C') if c.value<0 else ('ECEFF1','555555')
            else:
                fill,color='FFF2CC','7D5A00'
            c.fill=PatternFill('solid',fgColor=fill)
            c.font=Font(name='Arial',size=10,bold=column==14,color=color)
        if row[13].value is None:
            for column in (5,15):
                c=ws.cell(row[0].row,column)
                c.fill=PatternFill('solid',fgColor='FFF2CC')
                c.font=Font(name='Arial',size=10,color='7D5A00')
        ws.row_dimensions[row[0].row].height=44
    for col,width in {'A':18,'B':19,'C':10,'D':18,'E':28,'F':14,'G':18,'H':20,
                      'I':18,'J':18,'K':18,'L':22,'M':18,'N':22,'O':42,'P':27}.items():
        ws.column_dimensions[col].width=width

def add_order_economics(wb, orders, health):
    from .finance_report import table
    missing=sum(r['after_reserve'] is None for r in orders)
    note=(f'Денежный вклад по заказам: платежи покупателя + денежные компенсации − услуги − себестоимость. '
        f'Баллы показаны отдельно и не прибавлены к деньгам. Резерв {TAX_RATE:.0%} — допущение из настроек, не налоговая декларация. '
        f'Общие расходы бизнеса не включены. По {missing} позициям расчёт отсутствует (возвраты, незавершённые заказы или нет данных); '
        'они не считаются нулевой прибылью. Данные могут уточняться Маркетом. '+health)
    cols=['sku','order','model','created','status','qty','buyer','other','bonus','fees','cost','contribution','tax_reserve','after_reserve','quality','allocation']
    table(wb,'Экономика заказов',note,
        ['SKU','Заказ','Модель','Дата оформления','Статус товара','Доставлено, шт.','Покупатель, ₽',
         'Денежные компенсации, ₽','Баллы отдельно','Услуги, ₽','Себестоимость, ₽',
         'Денежный вклад до резерва, ₽',f'Резерв {TAX_RATE:.0%}, ₽','Вклад после резерва, ₽','Полнота','Распределение услуг'],
        [[r[c] for c in cols] for r in orders],6)
    style_order_economics(wb['Экономика заказов'])
    agg=defaultdict(lambda:{'known':0,'missing':0,'value':0.0})
    for r in orders:
        a=agg[(r['created'][:7],r['sku'])]
        if r['after_reserve'] is None: a['missing']+=1
        else: a['known']+=1;a['value']+=r['after_reserve']
    table(wb,'Вклад заказов по месяцам',
        'Группировка по месяцу оформления заказа. Сумма только рассчитанных позиций; '
        'при наличии неполных позиций это частичный результат. Граничные месяцы окна выгрузки неполные.',
        ['Месяц','SKU','Рассчитано позиций','Не рассчитано позиций','Вклад после резерва, ₽'],
        [[m,sku,a['known'],a['missing'],round(a['value'],2) if a['known'] else None] for (m,sku),a in sorted(agg.items())],5)
