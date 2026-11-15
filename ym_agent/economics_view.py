"""One decision view; cash reconciliation and closed months remain separate."""
from copy import copy
from pathlib import Path
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter
from .sku_diagnostics import summarize_skus


def add_economics_view(wb, orders, records, date_from, date_to, closed_dir=None):
    data=summarize_skus(orders,records,date_from,date_to)
    old=wb['Юнит-экономика'] if 'Юнит-экономика' in wb.sheetnames else None
    meta={}
    if old:
        for row in old.iter_rows(min_row=3,values_only=True):
            if len(row)>36 and row[1] in data:meta.setdefault(row[1],row)
        old.title='Модель по комиссиям';old.sheet_state='hidden'
        old['A1']='АРХИВ ПРЕЖНЕЙ МОДЕЛИ · для решений используйте новую Юнит-экономику'
    if 'Плюс-минус по товарам' in wb.sheetnames:
        wb['Плюс-минус по товарам'].sheet_state='hidden'
    ws=wb.create_sheet('Юнит-экономика');ws.sheet_properties.tabColor='287D6B';ws.sheet_view.showGridLines=False
    ws.merge_cells('A1:J1');ws['A1']='ЮНИТ-ЭКОНОМИКА · предварительный расчёт на штуку';ws.row_dimensions[1].height=30
    ws['A1'].font=Font(size=16,bold=True,color='23465A')
    ws.merge_cells('A2:J2');ws['A2']=f'База: заказы {date_from}–{date_to}. D — вклад доставок без невыкупов и возвратов. G — сценарий цены F при прежней доле поступлений и расходов. Не факт закрытого месяца.'
    ws.merge_cells('A3:J3');ws['A3']='Учтены закуп, налог 7%, услуги (включая буст) и связанное с заказами сальдо баллов. Возвраты с незавершённым расчётом и общие расходы бизнеса не включены. Детали раскрываются кнопкой + справа.'
    ws.merge_cells('A4:J4');ws['A4']='Жёлтая F — цена для проверки сценария. Баллы не равны деньгам; их распределение по заказам может быть неполным. Ставки и буст за 14 дней справа — справочно, повторно не вычитаются.'
    for ri in (2,3,4):ws.row_dimensions[ri].height=33;ws.cell(ri,1).alignment=Alignment(wrap_text=True,vertical='center')
    heads=['SKU','Товар','Доставлено, шт.','База без возвратов, ₽/шт.','Текущая цена, ₽','Цена сценария, ₽','Сценарий, ₽/шт.','Цена для 100 ₽/шт.','Поднять на, ₽','Ограничения расчёта',
      'Вклад с невыкупами, ₽/шт.','Невыкупы / отмены, поз.','Возвраты, поз.','Не рассчитано, поз.','Без проводок, поз.',
      'Поступления доставок, ₽/шт.','Услуги деньгами, ₽/шт.','Закуп, ₽/шт.','Налог доставок, ₽/шт.','Баллы доставок: нетто, ₽/шт.',
      'Вклад доставок деньгами, ₽/шт.','Вклад всех рассчитанных, ₽','Невыкупы / отмены: вклад, ₽','Баллы начислены, ₽','Баллы списаны, ₽',
      'Ставка буста сейчас, %','Буст NET, 14д, ₽/шт.','Буст GROSS, 14д, ₽/шт.','Канал (справочно)','Основание канала']
    ws.append(heads)
    prices={}
    if 'Индекс цен' in wb.sheetnames:
        for row in wb['Индекс цен'].values:
            if len(row)>2 and isinstance(row[2],(int,float)):prices[row[1]]=row[2]
    for sku,a in sorted(data.items(),key=lambda kv:kv[1]['adjusted']/kv[1]['qty'] if kv[1]['qty'] else -1e20):
        q=a['qty'];m=meta.get(sku,());num=lambda key:round(a[key]/q,4) if q else None
        status=[]
        if not q:status.append('Нет рассчитанных доставок')
        elif q<10:status.append('Мало данных: <10 шт.')
        if a['delivered_unmatched']:status.append('Нет проводок доставок: цену не рассчитываем')
        if a['missing']:status.append(f"Не рассчитано {int(a['missing'])} поз., включая возвраты")
        if not prices.get(sku):status.append('Нет текущей цены')
        if q and a['delivered_adjusted']+a['cost']<=0:status.append('Цена не выводится: сначала снизить расходы')
        if not status:status.append('Предварительно; скидки и услуги могут уточняться')
        tax=(a['buyer']-a['fees']-a['cost']-a['delivered_cash'])/q if q else None
        vals=[sku,m[0] if m else sku,q,num('delivered_adjusted'),prices.get(sku),prices.get(sku),None,None,None,'; '.join(status),num('adjusted'),a['cancelled'],a['returned'],a['missing'],a['unmatched'],num('buyer'),num('fees'),num('cost'),round(tax,4) if tax is not None else None,round((a['delivered_adjusted']-a['delivered_cash'])/q,4) if q else None,num('delivered_cash'),round(a['adjusted'],2) if a['known'] else None,round(a['cancel_adjusted'],2),round(a['credit'],2),round(a['spent'],2),m[30] if m else None,m[33] if m else None,m[34] if m else None,m[35] if m else None,m[36] if m else None]
        ws.append(vals);r=ws.max_row
        if q and prices.get(sku) and not a['delivered_unmatched']:
            ws.cell(r,7,f'=IF(AND(ISNUMBER(F{r}),F{r}>0),(D{r}+R{r})*F{r}/E{r}-R{r},"")')
            ws.cell(r,8,f'=IF(D{r}>=100,E{r},IF(D{r}+R{r}>0,ROUNDUP(E{r}*(R{r}+100)/(D{r}+R{r})/10,0)*10,""))')
            ws.cell(r,9,f'=IF(ISNUMBER(H{r}),MAX(0,H{r}-E{r}),"")')
        for c in ws[r]:
            c.font=Font(name='Arial',size=10,color='333333');c.alignment=Alignment(vertical='center',wrap_text=c.column in (2,10,30));c.number_format='#,##0.00;[Red](#,##0.00);"—"' if c.column>=4 else '0' if c.column==3 else 'General'
        for ci in (12,13,14,15):ws.cell(r,ci).number_format='0'
        ws.cell(r,6).fill=PatternFill('solid',fgColor='FFF2CC');ws.cell(r,6).font=Font(color='0000FF',size=10)
        ws.row_dimensions[r].height=48
    for c in ws[5]:c.fill=PatternFill('solid',fgColor='23465A');c.font=Font(bold=True,color='FFFFFF',size=10);c.alignment=Alignment(wrap_text=True,vertical='center',horizontal='center')
    ws.row_dimensions[5].height=60
    widths={'A':17,'B':30,'C':12,'D':17,'E':15,'F':15,'G':17,'H':17,'I':15,'J':39}
    for i in range(1,31):ws.column_dimensions[get_column_letter(i)].width=widths.get(get_column_letter(i),19)
    ws.column_dimensions['AD'].width=48
    for start,end in [('K','O'),('P','Y'),('Z','AD')]:ws.column_dimensions.group(start,end,hidden=True)
    ws.sheet_properties.outlinePr.summaryRight=False;ws.freeze_panes='C6';ws.auto_filter.ref=f'A5:AD{ws.max_row}'
    for col in ('D','G','K'):
        ws.conditional_formatting.add(f'{col}6:{col}{ws.max_row}',CellIsRule(operator='lessThan',formula=['0'],fill=PatternFill('solid',fgColor='FBE3E2'),font=Font(color='A61C1C')))
        ws.conditional_formatting.add(f'{col}6:{col}{ws.max_row}',CellIsRule(operator='greaterThanOrEqual',formula=['0'],fill=PatternFill('solid',fgColor='DFF0E4'),font=Font(color='1B5E20')))
    end=ws.max_row
    notes=[
      'Что объединено: диагностика плюса/минуса, расчёт без невыкупов, закуп и налог, цена для 100 ₽/шт., сценарий вашей цены, справочные буст и канал.',
      'D = (вклад доставленных заказов деньгами + начисленные баллы − списанные баллы по этим заказам) / доставленные штуки. Это предварительная экономика, не выплата.',
      'G = (D + закуп/шт.) × цена сценария / текущая цена − закуп/шт. Предположение: прежние доли выручки, услуг, налога и баллов; текущая цена служит базой сценария, не исторической ценой всех заказов.',
      'K учитывает известный вклад отмен и невыкупов. Возвраты и незавершённые позиции, не рассчитанные источником, в прибыль не включены; их количество видно в M/N.',
      'Не перенесены как показатели прибыли: деньги из Сверки выплат, взвешивание окон 7/30/90, сравнение с Озоном без сопоставимого периода. Прежняя модель комиссий сохранена скрытой.',
      'Сверка выплат: только денежные проводки по статусу и дате. Не содержит закупа, налога и прибыли. Прочее может быть ненулевым: например, компенсация утраты.',
      'Сентябрь 2026: отдельный календарный месяц, услуги сверены с актом. Закуп из справочника; пригодность возвратов к повторной продаже — допущение. Это не прогноз следующей продажи.',
    ]
    for i,note in enumerate(notes,end+3):
        ws.merge_cells(start_row=i,start_column=1,end_row=i,end_column=10);ws.cell(i,1,note);ws.cell(i,1).alignment=Alignment(wrap_text=True,vertical='center');ws.row_dimensions[i].height=36
    if 'Сверка выплат' in wb.sheetnames:
        cash=wb['Сверка выплат'];cash['A1']=f'ДЕНЕЖНЫЕ ПРОВОДКИ · {date_from}–{date_to}. D = G + H + I + J (расходы отрицательные). Только проведённые денежные статусы. E/F — отправка и ожидание, K/L — справочные баллы; в D не входят. Это не прибыль и не банковская выписка. Закуп, налог и прибыль здесь НЕ рассчитываются. Строка Σ период — сумма FBS + FBY за тот же период, не 90-дневная прибыль. Детали магазинов свёрнуты; раскрывайте +. Для цены и расчёта на штуку — «Юнит-экономика»; для закрытого месяца — отдельный лист.'
        cash.row_dimensions[1].height=95
        for row in cash.iter_rows(min_row=3):
            if row[2].value in ('FBS','FBY'):cash.row_dimensions[row[0].row].hidden=True;cash.row_dimensions[row[0].row].outlineLevel=1
        cash.sheet_properties.outlinePr.summaryBelow=True
    month_tabs = []
    if closed_dir:
        for path in sorted(Path(closed_dir).glob('*.xlsx')):
            if path.name.startswith('~$'):continue
            title = _add_closed(wb,path)
            if title:
                month_tabs.append(title)
    from .recent_economics import add_recent_economics
    ws=add_recent_economics(wb,orders,records,date_from,date_to,meta,prices)
    from .price_targets_returns import add_price_targets_returns
    add_price_targets_returns(wb,orders,records,date_from,date_to,meta,prices)
    fronts=['Дашборд','Юнит-экономика','Цены по товарам']+sorted(month_tabs, key=lambda t: (int(t.split()[1]), MONTHS.index(t.split()[0])), reverse=True)
    helpers=['Сверка выплат','Сентябрь — итог','Экономика заказов','Финансовые операции','Модель по комиссиям','Плюс-минус по товарам','Сентябрь — продажи','Сентябрь — услуги']
    wb._sheets=[wb[n] for n in fronts if n in wb.sheetnames]+[s for s in wb if s.title not in fronts+helpers]+[wb[n] for n in helpers if n in wb.sheetnames]
    return ws


MONTHS = 'Январь Февраль Март Апрель Май Июнь Июль Август Сентябрь Октябрь Ноябрь Декабрь'.split()


def _add_closed(wb,path):
    import openpyxl
    source=openpyxl.load_workbook(path)
    summary = next((name for name in source.sheetnames if name.endswith(' — итог')), None)
    if not summary:
        return
    label = summary.split(' — ')[0]
    title = str(source[summary]['A2'].value).split(' · ')[0]
    if label not in MONTHS or title.split()[:1] != [label]:
        return
    mapping = {summary: summary, 'Прибыль по артикулам': title,
               'Реализация — источник': label + ' — продажи', 'Услуги — источник': label + ' — услуги'}
    if not all(n in source.sheetnames for n in mapping):return
    for name,new in mapping.items():
        if new in wb.sheetnames:del wb[new]
        src=source[name];dst=wb.create_sheet(new)
        rows = list(src)
        if new == title:
            rows[5:] = sorted(rows[5:], key=lambda row: str(row[0].value).casefold())
        for dest_row, row in enumerate(rows, 1):
            for cell in row:
                val=cell.value
                if cell.data_type=='f':
                    for before,after in mapping.items():val=val.replace("'"+before+"'!","'"+after+"'!")
                if cell.data_type == 'f' and dest_row != cell.row:
                    from openpyxl.formula.translate import Translator
                    val = Translator(val, origin=cell.coordinate).translate_formula(row_delta=dest_row-cell.row)
                out=dst.cell(dest_row,cell.column,val)
                if cell.has_style:
                    out.font=copy(cell.font);out.fill=copy(cell.fill);out.border=copy(cell.border)
                    out.alignment=copy(cell.alignment);out.number_format=cell.number_format;out.protection=copy(cell.protection)
                if cell.hyperlink:out.hyperlink=copy(cell.hyperlink)
        for key,dim in src.column_dimensions.items():dst.column_dimensions[key]=copy(dim)
        for key,dim in src.row_dimensions.items():dst.row_dimensions[key]=copy(dim)
        for merged in src.merged_cells.ranges:dst.merge_cells(str(merged))
        dst.freeze_panes=src.freeze_panes;dst.sheet_view.showGridLines=False;dst.conditional_formatting=copy(src.conditional_formatting)
        if name in ('Реализация — источник','Услуги — источник'):dst.sheet_state='hidden'
    month = wb[title]
    month['A1'] = 'Итого за месяц, ₽'
    month['C1'] = f"='{summary}'!B16"
    month['C1'].number_format = '#,##0.00'
    month['J1'] = 'Сверка итогов →'
    month['J1'].hyperlink = f"#'{summary}'!A8"
    month['A2'] = title.upper() + ' · реализация и услуги по акту'
    note = 'Без накладных расходов; компенсации учитываются отдельно.' if title == 'Сентябрь 2026' else 'Без накладных расходов.'
    month['A4'] = 'После налога 7%; закуп по справочнику, возвраты пригодны к продаже. ' + note
    month.auto_filter.ref = f'A5:N{month.max_row}'
    month.sheet_properties.tabColor = '2F6B48'
    bonus_path = path.with_suffix(".bonus.json")
    if bonus_path.is_file():
        import json
        _add_monthly_bonus(month, json.loads(bonus_path.read_text(encoding="utf-8")))
    delivery_path = path.with_suffix(".delivery.json")
    if delivery_path.is_file():
        import json
        audit_name = label + ' — заказы'
        if audit_name in wb.sheetnames:
            del wb[audit_name]
        _add_delivery_month(wb, month, label, json.loads(delivery_path.read_text(encoding="utf-8")), summary)
    source.close()
    return title


def _add_monthly_bonus(month, balances):
    """Add monthly settled bonus turnover, including reversals, to the closed result."""
    from openpyxl.formatting.rule import CellIsRule

    last = month.max_row
    # M contains the cost referenced by monthly formulas. Move only the display
    # column K (return fees) so that J and K can be compared side by side.
    month.move_range(f'K5:K{last}', cols=4)
    headers = {
        'K': 'Прибыль с баллами, ₽/чистую шт.',
        'P': 'Баллы за скидки, с возвратами, ₽',
        'Q': 'Оплата услуг баллами, с возвратами, ₽',
        'R': 'Сальдо баллов, ₽',
    }
    for column, title in headers.items():
        cell = month[f'{column}5']
        cell._style = copy(month['J5']._style)
        cell.value = title
        month.column_dimensions[column].width = 21
    month.column_dimensions['O'].width = 20
    for row in range(6, last + 1):
        sku = month[f'A{row}'].value
        credit, service = balances[sku]
        for column, value in {
            'P': credit, 'Q': service, 'R': f'=P{row}+Q{row}',
            'K': f'=IF(B{row}>C{row},(I{row}+R{row})/(B{row}-C{row}),"")'
        }.items():
            cell = month[f'{column}{row}']
            cell._style = copy(month[f'J{row}']._style)
            cell.value = value
    for op, color in [('lessThan', 'FCE5E5'), ('greaterThanOrEqual', 'E4F2E6')]:
        month.conditional_formatting.add(f'K6:K{last}', CellIsRule(
            operator=op, formula=['0'], fill=PatternFill('solid', fgColor=color)))
    month['A3'] = 'J — результат по акту. K — с сальдо баллов (P:R) по дате проводки за месяц, на выкупленную штуку. Невыкупы и возвраты включены.'
    month.auto_filter.ref = f'A5:R{last}'


def _add_delivery_month(wb, month, label, snapshot, summary):
    """Group settled transactions by the original delivered order, scheme and SKU."""
    audit = wb.create_sheet(label + ' — заказы')
    audit.append(['Артикул', 'Схема', 'Заказ', 'Доставка', 'Доставлено', 'Возвращено',
                  'По реализации, ₽', 'Оплата, ₽', 'Возврат оплаты, ₽', 'Услуги деньгами, ₽',
                  'Сальдо баллов, ₽', 'Закуп/шт., ₽', 'Неполные данные (1)', 'Результат, ₽'])
    for row in snapshot['rows']:
        audit.append(row)
        i = audit.max_row
        audit[f'N{i}'] = (f'=IF(M{i}=1,"",SUM(H{i}:K{i})-(H{i}+I{i})*\'{summary}\'!$B$4'
                          f'-(E{i}-F{i}*\'{summary}\'!$B$5)*L{i})')
    audit.freeze_panes = 'D2'
    audit.auto_filter.ref = audit.dimensions
    for cell in audit[1]:
        cell._style = copy(month['A5']._style)
        audit.column_dimensions[cell.column_letter].width = 21
    audit.row_dimensions[1].height = 48
    # Keep the original volume in S, including its use in the calendar-month result.
    last = month.max_row
    month.move_range(f'B5:B{last}', cols=17)
    for i in range(6, last + 1):
        for col in ('G', 'J', 'K'):
            cell = month[f'{col}{i}']
            cell.value = cell.value.replace(f'B{i}', f'S{i}')
    avg = wb[summary]['B17']
    avg.value = avg.value.replace(f"'{month.title}'!B6:B", f"'{month.title}'!S6:S")
    new = {'B':'По доставке с баллами, ₽/выкуп', 'T':'Доставлено, шт.', 'U':'Возвращено по этим заказам, шт.',
           'V':'Сальдо баллов этих заказов, ₽', 'W':'Прибыль по доставкам, ₽', 'X':'Заказных позиций без сверки'}
    for col, title in new.items():
        cell = month[f'{col}5']
        cell.value, cell._style = title, copy(month['J5']._style)
        month.column_dimensions[col].width = 22
    end = audit.max_row
    prefix = f"'{audit.title}'!"
    for i in range(6, month.max_row + 1):
        def total(col):
            return f'SUMIF({prefix}$A$2:$A${end},A{i},{prefix}${col}$2:${col}${end})'
        for col, src in [('T','E'),('U','F'),('V','K'),('X','M')]:
            month[f'{col}{i}'] = '=' + total(src)
        month[f'W{i}'] = f'=IF(OR(X{i}>0,T{i}=0),"",{total("N")})'
        month[f'B{i}'] = f'=IF(AND(X{i}=0,T{i}>U{i}),W{i}/(T{i}-U{i}),"")'
        for col in ('B','V','W'):
            month[f'{col}{i}']._style = copy(month[f'J{i}']._style)
    month['S1'] = 'Заказы и проводки →'
    month['S1'].hyperlink = f"#'{audit.title}'!A1"
    month['S2'] = 'По доставкам · проводки по ' + snapshot['through']
    month['S3'] = 'B = результат доставленных заказов после возвратов / выкупленные штуки. При неполной сверке пусто. Невыкупы без доставки и общие расходы вне расчёта.'
    month.auto_filter.ref = f'A5:X{month.max_row}'

    from openpyxl.formatting.rule import FormulaRule
    for condition, fill, font in [
        ('B6<50', 'FBE3E2', 'A61C1C'),
        ('AND(B6>=50,B6<=100)', 'FFF2CC', '7D5A00'),
        ('B6>100', 'DFF0E4', '1B5E20'),
    ]:
        month.conditional_formatting.add(f'B6:B{last}', FormulaRule(
            formula=[f'AND(ISNUMBER(B6),{condition})'],
            fill=PatternFill('solid', fgColor=fill), font=Font(color=font, bold=True)))
