"""Separate money, bonus, and order economics views; no cash-window profit."""
from collections import defaultdict
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from .finance_data import summarize

LEDGERS = ['cash_done','cash_sent','cash_pending','bonus_done','bonus_pending','excluded','unknown']
HEADERS = ['Проведено, ₽','В отправке, ₽','Ожидается, ₽','Баллы проведено, ₽',
           'Баллы ожидается, ₽','Исключено, ₽','Статус не распознан, ₽']

def add_settlement_sheet(wb, records, date_from, date_to, health):
    """Restore the SKU/model reconciliation layout with status-aware amounts."""
    name='Сверка выплат'
    index=wb.sheetnames.index(name) if name in wb.sheetnames else len(wb.sheetnames)
    if name in wb.sheetnames: del wb[name]
    ws=wb.create_sheet(name,index)
    ws.sheet_properties.tabColor='C0392B'
    ws.sheet_view.showGridLines=False
    names={}
    if 'Юнит-экономика' in wb.sheetnames:
        for row in wb['Юнит-экономика'].iter_rows(values_only=True):
            if len(row)>1 and isinstance(row[0],str) and isinstance(row[1],str): names[row[1]]=row[0]
    selected=[r for r in records if str(date_from)<=r['date']<=str(date_to)]
    summary=summarize(selected,date_from,date_to)
    components=defaultdict(lambda:defaultdict(float))
    for r in selected:
        if r['ledger']!='cash_done': continue
        key={'Платёж покупателя':'buyer','Оплата услуг Маркета':'services',
             'Возврат платежа покупателя':'refunds'}.get(r['source'],'other')
        components[(r['sku'],r['model'])][key]+=r['amount']
    headers=['Товар','SKU','Модель','Проведено, ₽','В отправке, ₽','Ожидается, ₽',
             'Платежи покупателей, ₽','Удержано за услуги, ₽','Возвраты денег, ₽','Прочие деньги, ₽',
             'Баллы: сальдо, ₽','Баллы ожидаются, ₽','Исключено, ₽','Неизвестный статус, ₽']
    ws.merge_cells('A1:N1')
    ws['A1']=(f'СВЕРКА ВЫПЛАТ · {date_from}–{date_to}. {health}. '
        'Под каждым SKU — итог FBS + FBY за тот же период (не дополнительные операции). '
        'Проведено = платежи + услуги + возвраты + прочие деньги; расходы со знаком минус. '
        'Отправка, ожидание, баллы и исключённые операции учитываются отдельно. '
        'Это движение денег по дате транзакции, не прибыль и не банковская выписка. '
        'Доходы и расходы одного заказа могут попадать в разные периоды; вклад заказов — на листе «Экономика заказов».')
    ws['A1'].font=Font(name='Arial',size=10,color='555555')
    ws['A1'].alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[1].height=72
    ws.append(headers)
    head_colors=['23465A']*3+['365C85','365C85','365C85','287D6B','A96325','A96325','596773','7954A0','7954A0','596773','9A7618']
    for c in ws[2]:
        c.fill=PatternFill('solid',fgColor=head_colors[c.column-1])
        c.font=Font(name='Arial',size=10,bold=True,color='FFFFFF')
        c.alignment=Alignment(wrap_text=True,horizontal='center',vertical='center')
    ws.row_dimensions[2].height=48
    sku_totals=defaultdict(float)
    for (sku,model),a in summary.items(): sku_totals[sku]+=a['cash_done']
    all_values=[];kinds={}
    for group_no,sku in enumerate(sorted(sku_totals,key=lambda s:(-sku_totals[s],s))):
        sku_values=[]
        for key in sorted(k for k in summary if k[0]==sku):
            a=summary[key];b=components[key]
            values=[a['cash_done'],a['cash_sent'],a['cash_pending']]+[round(b[k],2) for k in ['buyer','services','refunds','other']]+[a[k] for k in ['bonus_done','bonus_pending','excluded','unknown']]
            ws.append([names.get(sku,sku)[:50],sku,key[1]]+values)
            kinds[ws.max_row]=('detail',group_no)
            sku_values.append(values);all_values.append(values)
        ws.append(['Итого по артикулу',sku,'Σ период']+[round(sum(v[j] for v in sku_values),2) for j in range(11)])
        kinds[ws.max_row]=('subtotal',group_no)
    ws.append(['ИТОГО' if selected else 'НЕТ ДАННЫХ','','']+([round(sum(v[j] for v in all_values),2) for j in range(11)] if selected else [None]*11))
    kinds[ws.max_row]=('total',0)
    line=Side(style='thin',color='DDDDDD')
    for row in ws.iter_rows(min_row=3):
        kind,group_no=kinds[row[0].row]
        for c in row:
            bg='EEEEEE' if kind=='total' else 'FFF0DA' if kind=='subtotal' else 'F3F6F8' if group_no%2 else 'FFFFFF'
            color='333333'
            if isinstance(c.value,(int,float)):
                color='1B5E20' if c.value>0 else 'B71C1C' if c.value<0 else '999999'
                c.number_format='#,##0.00;[Red]-#,##0.00;"—"'
            if kind=='detail':
                if c.column==7: bg='E8F4EF'
                elif c.column in (8,9): bg='FFF0E0'
                elif c.column in (11,12): bg='F1EAF7'
            if c.column==4 and isinstance(c.value,(int,float)):
                bg='DFF0E4' if c.value>0 else 'FBE3E2' if c.value<0 else 'ECEFF1'
            if c.column==14 and c.value: bg='FFF2CC';color='7D5A00'
            c.fill=PatternFill('solid',fgColor=bg)
            c.font=Font(name='Arial',size=10,bold=kind!='detail' or c.column==4,color=color)
            c.border=Border(bottom=line)
            c.alignment=Alignment(horizontal='left' if c.column<3 else 'center',vertical='center')
        ws.row_dimensions[row[0].row].height=24
    ws.column_dimensions['A'].width=40;ws.column_dimensions['B'].width=16;ws.column_dimensions['C'].width=13
    for i in range(4,15): ws.column_dimensions[get_column_letter(i)].width=19
    ws.freeze_panes='G3'
    ws.auto_filter.ref=f'A2:N{max(2,ws.max_row-1)}'
    return ws

def add_monthly_money(wb, records, date_from, date_to):
    """SKU/month matrix in the former monthly-profit report's visual style."""
    name = 'Деньги по месяцам'
    index = wb.sheetnames.index(name) if name in wb.sheetnames else len(wb.sheetnames)
    if name in wb.sheetnames: del wb[name]
    ws = wb.create_sheet(name, index)
    ws.sheet_properties.tabColor = '8E44AD'
    ws.sheet_view.showGridLines = False
    selected = [r for r in records if str(date_from) <= r['date'] <= str(date_to)]
    months = sorted({r['date'][:7] for r in selected})
    amounts = defaultdict(lambda: defaultdict(lambda: defaultdict(float)))
    names = {}
    if 'Юнит-экономика' in wb.sheetnames:
        for row in wb['Юнит-экономика'].iter_rows(values_only=True):
            if len(row)>1 and isinstance(row[0],str) and isinstance(row[1],str):
                names[row[1]] = row[0]
    for r in selected:
        amounts[r['sku']][r['date'][:7]][r['ledger']] += r['amount']
    width = 2 + 7 * (len(months)+1)
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=width)
    ws['A1'] = (f'ДЕНЬГИ ПО МЕСЯЦАМ · {date_from}–{date_to}. Операции по SKU, FBS + FBY, по дате транзакции. '
        'Это движение денег, не прибыль. Граничные месяцы неполные. '
        'В каждом месяце: проведено / в отправке / ожидается. Баллы, исключённые и неизвестные статусы — '
        'в свёрнутых колонках (раскройте «+» сверху). Пусто — операций нет; 0 — нулевое сальдо. '
        'Сортировка по проведённым деньгам за весь период ↓.')
    ws['A1'].font = Font(name='Arial',size=10,color='555555')
    ws['A1'].alignment = Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[1].height = 80
    # Keep the note readable in the first visible screen instead of across the full matrix.
    ws.unmerge_cells(start_row=1,start_column=1,end_row=1,end_column=width)
    ws.merge_cells('A1:P1')
    tints = ['FDF3E2','E8F1F8','EDE7F6','F2EDF8','F2EDF8','EEEEEE','FFF2CC']
    line = Side(style='thin',color='DDDDDD')
    for c,label in [(1,'Товар'),(2,'SKU')]:
        ws.cell(2,c,label); ws.merge_cells(start_row=2,start_column=c,end_row=3,end_column=c)
    groups = months + ['ИТОГО ЗА ПЕРИОД']
    for i,m in enumerate(groups):
        base = 3 + 7*i
        title = m
        if i<len(months) and (m==str(date_from)[:7] or m==str(date_to)[:7]):
            title += ' · в пределах окна'
        ws.cell(2,base,title)
        ws.merge_cells(start_row=2,start_column=base,end_row=2,end_column=base+6)
        for j,h in enumerate(HEADERS):
            ws.cell(3,base+j,h)
            ws.column_dimensions[get_column_letter(base+j)].width = 15
        ws.column_dimensions.group(get_column_letter(base+3),get_column_letter(base+6),hidden=True)
    for row in ws.iter_rows(min_row=2,max_row=3,max_col=width):
        for c in row:
            c.fill = PatternFill('solid',fgColor='8E44AD')
            c.font = Font(name='Arial',size=10,bold=True,color='FFFFFF')
            c.alignment = Alignment(horizontal='center',vertical='center',wrap_text=True)
    for i in range(len(groups)):
        for j,tint in enumerate(tints):
            c=ws.cell(3,3+7*i+j)
            c.fill=PatternFill('solid',fgColor=tint)
            c.font=Font(name='Arial',size=10,bold=True,color='333333')
    ws.row_dimensions[2].height=24; ws.row_dimensions[3].height=32
    skus=sorted(amounts,key=lambda sku:(-sum(v['cash_done'] for v in amounts[sku].values()),sku))
    for ri,sku in enumerate(skus,4):
        ws.cell(ri,1,names.get(sku,sku)[:50]);ws.cell(ri,2,sku)
        for i,m in enumerate(groups):
            values=amounts[sku].get(m) if i<len(months) else {
                k:sum(v[k] for v in amounts[sku].values()) for k in LEDGERS}
            for j,k in enumerate(LEDGERS):
                if values is not None: ws.cell(ri,3+7*i+j,round(values[k],2))
    total_row=4+len(skus)
    ws.cell(total_row,1,'ИТОГО' if selected else 'НЕТ ДАННЫХ')
    for c in range(3,width+1):
        if selected: ws.cell(total_row,c,round(sum(ws.cell(r,c).value or 0 for r in range(4,total_row)),2))
    for row in ws.iter_rows(min_row=4,max_row=total_row,max_col=width):
        for c in row:
            total=c.row==total_row
            c.border=Border(left=line,right=line,top=line,bottom=line)
            c.alignment=Alignment(horizontal='left' if c.column==1 else 'center',vertical='center')
            color='333333'
            if isinstance(c.value,(float,int)):
                c.number_format='#,##0.00;[Red]-#,##0.00;"—"'
                color='1B5E20' if c.value>0 else 'B71C1C' if c.value<0 else '999999'
            c.font=Font(name='Arial',size=10,bold=total or c.column>=3+7*len(months),color=color)
            if total: c.fill=PatternFill('solid',fgColor='EEEEEE')
            elif c.column>=3: c.fill=PatternFill('solid',fgColor=tints[(c.column-3)%7])
        ws.row_dimensions[row[0].row].height=22
    ws.column_dimensions['A'].width=40;ws.column_dimensions['B'].width=16
    ws.freeze_panes='C4'
    ws.auto_filter.ref=f'A3:{get_column_letter(width)}{max(3,total_row-1)}'
    return ws

def table(wb, name, note, headers, rows, money_start=3):
    if name in wb.sheetnames: del wb[name]
    ws = wb.create_sheet(name)
    ws.append([note]); ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(headers))
    ws['A1'].alignment = Alignment(wrap_text=True, vertical='top')
    ws.row_dimensions[1].height = 66
    ws.append(headers)
    for row in rows: ws.append(row)
    ws.freeze_panes = 'C3'; ws.auto_filter.ref = f'A2:{get_column_letter(len(headers))}{max(2,ws.max_row)}'
    ws.sheet_view.showGridLines = False
    for c in ws[2]:
        c.font = Font(name='Arial',bold=True,color='FFFFFF')
        c.fill = PatternFill('solid',fgColor='23465A')
        c.alignment = Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[2].height = 40
    for i in range(1,len(headers)+1): ws.column_dimensions[get_column_letter(i)].width = 22 if i<3 else 19
    for row in ws.iter_rows(min_row=3):
        for c in row:
            c.font = Font(name='Arial',size=10)
            if c.column >= money_start and isinstance(c.value,(int,float)): c.number_format = '#,##0.00'
    return ws

def add_finance_sheets(wb, records, orders, date_from, date_to, health):
    for name in ['Сверка выплат','Прибыль по месяцам','Деньги по месяцам','Экономика заказов','Финансовые операции']:
        if name in wb.sheetnames: del wb[name]
    summary = summarize(records, date_from, date_to) if date_from and date_to else {}
    rows = [[sku,model]+[a[k] for k in LEDGERS] for (sku,model),a in sorted(summary.items())]
    totals = [round(sum(a[k] for a in summary.values()),2) for k in LEDGERS]
    rows.append(['ИТОГО' if records else 'НЕТ ДАННЫХ','']+(totals if records else [None]*len(LEDGERS)))
    note = (f'Операции {date_from}–{date_to}, отбор по дате транзакции. {health}. '
        'Деньги, отправленные/ожидаемые суммы и справочные баллы показаны отдельно. '
        'Отменённые и неизвестные статусы не включены в деньги. Это не прибыль и не банковская выписка: '
        'оплата и услуги одного заказа могут относиться к разным периодам. Отсутствие продаж не обнуляет расходы.')
    add_settlement_sheet(wb,records,date_from,date_to,health)
    selected = [r for r in records if str(date_from)<=r['date']<=str(date_to)]
    add_monthly_money(wb, records, date_from, date_to)
    table(wb,'Финансовые операции',
        'Исходные операции после нормализации. Проверяйте статус, заказ и дату; справочные и отменённые записи сохранены для сверки.',
        ['SKU','Модель','Дата','Заказ','Источник','Услуга / товар','Статус','Учёт','Сумма, ₽'],
        [[r['sku'],r['model'],r['date'],r['order'],r['source'],r['service'],r['status'],r['ledger'],r['amount']] for r in selected],9)
    ws = wb['Финансовые операции']
    for col in ['E','F','G']: ws.column_dimensions[col].width=45
    from .order_economics import add_order_economics
    add_order_economics(wb, orders, health)
