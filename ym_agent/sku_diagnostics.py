"""Order-cohort profitability bridge, including separately booked bonus offsets."""
from collections import defaultdict

def price_target(current_price, delivered_profit, unit_cost, quantity):
    """Scenario: price scales receipts and service/reserve charges proportionally."""
    from math import ceil
    if not quantity or not current_price or current_price<=0: return None,None,'Нет цены или доставок'
    if delivered_profit>=100: return 0,current_price,'Цель уже достигнута'
    retained=delivered_profit+unit_cost
    if retained<=0: return None,None,'Модель не даёт цену: сначала снизить расходы'
    target=ceil(current_price*(100+unit_cost)/retained/10)*10
    return target-current_price,target,'Мало данных (<5 шт.)' if quantity<5 else 'Оценка при сохранении доли расходов'

def add_price_targets(wb,ws,data):
    from openpyxl.styles import Font,PatternFill,Alignment
    from openpyxl.utils import get_column_letter
    prices={}
    if 'Индекс цен' in wb.sheetnames:
        for row in wb['Индекс цен'].values:
            if len(row)>2 and isinstance(row[2],(int,float)):prices[str(row[1])]=row[2]
    # Add the decision columns beside the existing headline result.
    widths={i:ws.column_dimensions[get_column_letter(i)].width for i in range(1,23)}
    ws.unmerge_cells('A1:K1');ws.insert_cols(5,5);ws.merge_cells('A1:P1')
    for i,width in widths.items():ws.column_dimensions[get_column_letter(i if i<5 else i+5)].width=width
    labels=['Вклад без невыкупов / возвратов, ₽/шт.','Текущая цена, ₽','Поднять на, ₽','Цена для 100 ₽/шт., модель','Основание рекомендации']
    for i,label in enumerate(labels,5):
        c=ws.cell(2,i,label);c.fill=PatternFill('solid',fgColor='7954A0')
        c.font=Font(name='Arial',size=10,bold=True,color='FFFFFF')
        c.alignment=Alignment(wrap_text=True,vertical='center',horizontal='center')
        ws.column_dimensions[get_column_letter(i)].width=21 if i<9 else 42
    for ri in range(3,ws.max_row+1):
        sku=ws.cell(ri,2).value;a=data[sku];q=a['qty']
        profit=a['delivered_adjusted']/q if q else None
        cost=a['cost']/q if q else 0
        current=prices.get(sku)
        delta,target,status=price_target(current,profit,cost,q)
        if a['delivered_unmatched']:delta=target=None;status='Нет проводок доставок — сначала сверить баллы'
        elif a['missing']:status+='; есть незавершённые позиции'
        for ci,value in enumerate([round(profit,2) if profit is not None else None,current,delta,target,status],5):
            c=ws.cell(ri,ci,value);c.number_format='#,##0.00;[Red]-#,##0.00;"—"'
            c.fill=PatternFill('solid',fgColor='FFF2CC' if profit is None or profit<100 else 'DFF0E4')
            c.font=Font(name='Arial',size=10,bold=ci in (7,8),color='333333')
            c.alignment=Alignment(wrap_text=ci==9,vertical='center')
    ws['A1']=str(ws['A1'].value)+' ЦЕНА: цель 100 ₽/доставленную шт. без невыкупов/возвратов, при прежнем резерве. Модель: новая цена = текущая × (себестоимость/шт. + 100) / (себестоимость/шт. + вклад доставок/шт.), округление вверх до 10 ₽. Предполагается пропорциональный рост поступлений, услуг и резерва, неизменная себестоимость и структура скидок; это не гарантия прибыли. Текущая цена из листа «Индекс цен».'
    ws.row_dimensions[1].height=125;ws.row_dimensions[2].height=72
    ws.auto_filter.ref=f'A2:AA{ws.max_row}'

def summarize_skus(orders, transactions, date_from, date_to):
    bonuses=defaultdict(float); credits=defaultdict(float); matched=set()
    for t in transactions:
        if t['date']>str(date_to): continue
        key=(t['sku'],t['model'],t['order'])
        matched.add(key)
        if t['ledger']=='bonus_done':
            bonuses[key]+=t['amount']
            if str(t['source']).startswith(('Баллы за','Возврат баллов')):
                credits[key]+=t['amount']
    out={}
    for r in orders:
        if not str(date_from)<=r['created']<=str(date_to): continue
        a=out.setdefault(r['sku'],defaultdict(float))
        key=(r['sku'],r['model'],r['order'])
        returned=r['status'].startswith('Возврат')
        cancelled=r['status']=='Отменён' or r['status'].startswith('Невыкуп')
        if returned:
            a['returned']+=1; a['return_fees_known']+=r['fees'] or 0
        if r['after_reserve'] is None:
            a['missing']+=1
            continue
        a['known']+=1
        if key not in matched: a['unmatched']+=1
        a['cash']+=r['after_reserve']
        a['credit']+=credits[key]
        a['spent']+=credits[key]-bonuses[key]
        adjusted=r['after_reserve']+bonuses[key]
        a['adjusted']+=adjusted
        a['reserve']+=r['tax_reserve'] or 0
        if cancelled:
            a['cancelled']+=1
            a['cancel_fees']+=r['fees'] or 0
            a['cancel_adjusted']+=adjusted
        else:
            if key not in matched: a['delivered_unmatched']+=1
            a['qty']+=r['qty'] or 0
            a['buyer']+=(r['buyer'] or 0)+(r['other'] or 0)
            a['fees']+=r['fees'] or 0
            a['cost']+=r['cost'] or 0
            a['delivered_cash']+=r['after_reserve']
            a['delivered_adjusted']+=adjusted
    return out

def add_sku_diagnostics(wb,orders,transactions,date_from,date_to):
    from openpyxl.styles import Font,PatternFill,Alignment
    from openpyxl.utils import get_column_letter
    from .finance_report import table
    data=summarize_skus(orders,transactions,date_from,date_to)
    names={}
    if 'Юнит-экономика' in wb.sheetnames:
        for row in wb['Юнит-экономика'].values:
            if len(row)>1 and isinstance(row[0],str) and isinstance(row[1],str):names[row[1]]=row[0]
    headers=['Товар','SKU','Результат с баллами, ₽','На доставленную шт., ₽','Что влияет на результат',
        'Доставлено, шт.','Невыкупы / отмены, поз.','Возвраты, поз.','Не рассчитано, поз.',
        'Доставки с баллами, ₽','Невыкупы / отмены с баллами, ₽','Доставки деньгами, ₽',
        'Расходы невыкупов / отмен, ₽','Услуги возвратов, справочно ₽','Платежи доставок, ₽',
        'Услуги доставок деньгами, ₽','Себестоимость доставок, ₽','Резерв из расчёта, ₽',
        'Начислено баллов нетто, ₽','Оплачено баллами нетто, ₽','Результат без баллов, ₽','Полнота сверки']
    rows=[]
    for sku,a in sorted(data.items(),key=lambda kv:kv[1]['adjusted']):
        if not a['known']: reason='Нет рассчитанных заказов'
        elif a['adjusted']<0:
            reason='Невыкупы / отмены перевели в минус' if a['delivered_adjusted']>=0 else 'Доставки убыточны; проверить услуги и цену'
        else: reason='Плюс; невыкупы / отмены уменьшают результат' if a['cancel_adjusted']<0 else 'Предварительно в плюсе'
        quality=f"Не рассчитано: {int(a['missing'])}; без проводок: {int(a['unmatched'])}. Начисления могут уточняться"
        rows.append([names.get(sku,sku),sku,round(a['adjusted'],2) if a['known'] else None,
            round(a['adjusted']/a['qty'],2) if a['qty'] else None,reason,
            a['qty'],a['cancelled'],a['returned'],a['missing']]+[round(a[k],2) for k in
            ['delivered_adjusted','cancel_adjusted','delivered_cash','cancel_fees','return_fees_known','buyer','fees','cost','reserve','credit','spent','cash']]+[quality])
    ws=table(wb,'Плюс-минус по товарам',
        f'Заказы оформлены {date_from}–{date_to}. ПРЕДВАРИТЕЛЬНЫЙ вклад, не бухгалтерская чистая прибыль. '
        'С баллами = результат без баллов + начисленные баллы нетто − услуги за счёт баллов нетто. '
        'Проводки связаны по SKU, модели и заказу; премия отдельно не прибавляется повторно. '
        'Резерв сохранён из расчёта заказов; налог на возможную премию и общие расходы бизнеса не включены. '
        'Незавершённые возвраты не входят в результат; их известные услуги показаны справочно, не прибавляйте повторно. '
        'Балльные проводки могут быть неполными из-за окна выгрузки. Среднее = результат всех рассчитанных заказов / доставленные штуки.',headers,rows)
    ws.unmerge_cells('A1:V1'); ws.merge_cells('A1:K1')
    ws['A1'].font=Font(name='Arial',size=10,color='555555')
    ws.sheet_properties.tabColor='287D6B';ws.freeze_panes='C3'
    ws.row_dimensions[1].height=92;ws.row_dimensions[2].height=62
    for j in range(1,23):ws.column_dimensions[get_column_letter(j)].width=19
    for col,width in [('A',36),('B',16),('E',48),('V',55)]:ws.column_dimensions[col].width=width
    for row in ws.iter_rows(min_row=3):
        ws.row_dimensions[row[0].row].height=44
        for c in row:
            c.alignment=Alignment(wrap_text=c.column in (1,5,22),vertical='center')
            if 6<=c.column<=9:c.number_format='0'
            if c.column in (3,4,10,11,12,21) and isinstance(c.value,(int,float)):
                c.fill=PatternFill('solid',fgColor='DFF0E4' if c.value>=0 else 'FBE3E2')
                c.font=Font(name='Arial',size=10,bold=c.column in (3,4),color='1B5E20' if c.value>=0 else 'A61C1C')
            elif c.column in (19,20):c.fill=PatternFill('solid',fgColor='F1EAF7')
            elif c.column in (13,14,16,17,18):c.fill=PatternFill('solid',fgColor='FFF0E0')
            if c.column==9 and c.value:c.fill=PatternFill('solid',fgColor='FFF2CC')
    add_price_targets(wb,ws,data)
    return data
