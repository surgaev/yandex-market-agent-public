"""Recent order-cohort economics by SKU and fulfillment model."""
from datetime import date,timedelta
from collections import defaultdict
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.utils import get_column_letter
from .sku_diagnostics import summarize_skus


def metrics(a):
    q=a.get('qty',0)
    def pu(k):return a.get(k,0)/q if q else None
    return dict(q=q,profit=pu('delivered_adjusted'),cost=pu('cost'),buyer=pu('buyer'),fees=pu('fees'),
        tax=(a.get('buyer',0)-a.get('fees',0)-a.get('cost',0)-a.get('delivered_cash',0))/q if q else None,
        bonus=(a.get('delivered_adjusted',0)-a.get('delivered_cash',0))/q if q else None,
        missing=a.get('missing',0),unmatched=a.get('delivered_unmatched',0),cancel=pu('cancel_adjusted'))


def add_recent_economics(wb,orders,records,date_from,date_to,meta,prices):
    end=date.fromisoformat(str(date_to));start=date.fromisoformat(str(date_from))
    windows={'recent':(end-timedelta(days=13),end),'prior':(end-timedelta(days=27),end-timedelta(days=14)),
             'week':(end-timedelta(days=6),end),'all':(start,end)}
    summaries={}
    for model in ['Σ','FBS','FBY']:
        subset=orders if model=='Σ' else [r for r in orders if r['model']==model]
        for win,(lo,hi) in windows.items():
            cohort=[r for r in subset if str(max(start,lo))<=r['created']<=str(hi)]
            # Observe ALL cohorts through the same report cutoff, including later bonus postings.
            summaries[model,win]=summarize_skus(cohort,records,start,end)
    models=defaultdict(set)
    for r in orders:
        if str(start)<=r['created']<=str(end):models[r['sku']].add(r['model'])
    names={s:m[0] for s,m in meta.items()}
    for sheet in ['Товары','Плюс-минус по товарам','Индекс цен']:
        if sheet in wb.sheetnames:
            for row in wb[sheet].values:
                if len(row)>1 and row[1] in models and isinstance(row[0],str) and row[0]!=row[1]:names.setdefault(row[1],row[0])
    if 'Юнит-экономика' in wb.sheetnames:del wb['Юнит-экономика']
    ws=wb.create_sheet('Юнит-экономика');ws.sheet_view.showGridLines=False;ws.sheet_properties.tabColor='287D6B'
    title='ЮНИТ-ЭКОНОМИКА · динамика на штуку и FBS / FBY'
    subtitles=[title,
      f"14 дней: {windows['recent'][0]:%d.%m}–{end:%d.%m.%Y}; предыдущие: {windows['prior'][0]:%d.%m}–{windows['prior'][1]:%d.%m.%Y}. Отбор по дате оформления заказа, не по дате выплаты.",
      'Прибыль здесь — предварительный вклад доставленных заказов: закуп, налог 7%, услуги и сальдо баллов. Возвраты, невыкупы и общие расходы НЕ включены в показатель на штуку.',
      'Σ — средневзвешенный итог по штукам; FBS/FBY — отдельные расчёты. Пусто означает нет рассчитанных доставок. Цены и сценарий раскрываются через + над I:M; подробности — справа.'
    ]
    for ri,txt in enumerate(subtitles,1):ws.merge_cells(start_row=ri,start_column=1,end_row=ri,end_column=14);ws.cell(ri,1,txt);ws.cell(ri,1).alignment=Alignment(wrap_text=True,vertical='center');ws.row_dimensions[ri].height=32
    ws['A1'].font=Font(size=16,bold=True,color='23465A')
    heads=['SKU','Товар','Канал','Шт. за 14д','Последние 14д, ₽/шт.','Предыдущие 14д, ₽/шт.','Изменение, ₽/шт.','Весь период, ₽/шт.',
      'Текущая цена, ₽','Цена сценария, ₽','Сценарий, ₽/шт.','Цена для 100 ₽/шт.','Поднять на, ₽','Надёжность / сигнал',
      'Шт. предыдущие 14д','Шт. последние 7д','Последние 7д, ₽/шт.','Не рассчитано: 14д','Не рассчитано: пред.14д','Доставок без проводок: 14д','Доставок без проводок: пред.14д',
      'Закуп базы, ₽/шт.','Прибыль базы цены, ₽/шт.','Период базы цены','Изменение поступлений, ₽/шт.','Изменение услуг, ₽/шт.','Изменение закупа, ₽/шт.','Изменение налога, ₽/шт.','Изменение сальдо баллов, ₽/шт.',
      'Вклад невыкупов: 14д / доставл.шт.','Шт. весь период','Возвраты, поз. 14д','Ставка буста, справочно %','Буст NET 14д, справочно ₽/шт.','Буст GROSS 14д, справочно ₽/шт.']
    for ci,h in enumerate(heads,1):ws.cell(5,ci,h)
    for sku in sorted(models,key=lambda s:(s.casefold(),names.get(s,s).casefold())):
        channel_rows=sorted(models[sku]&{'FBS','FBY'});labels=['Σ']+channel_rows if len(channel_rows)>1 else channel_rows or ['Σ']
        first=ws.max_row+1
        for model in labels:
            a={win:metrics(summaries[model,win].get(sku,{})) for win in windows};x,y,z,allm=a['recent'],a['prior'],a['week'],a['all']
            delta=x['profit']-y['profit'] if x['q'] and y['q'] else None
            flags=[]
            if delta is None:flags.append('Нет базы сравнения')
            if x['q']<5 or y['q']<5:flags.append('Мало данных: <5 шт. в окне')
            if x['unmatched'] or y['unmatched']:flags.append('Неполные проводки — Δ не подтверждена')
            if x['missing'] or y['missing']:flags.append('Есть незавершённые позиции')
            if not flags:flags.append('Предварительно: снижение' if delta<0 else 'Предварительно: рост' if delta>0 else 'Без изменения')
            if delta is not None and delta<0 and (x['unmatched'] or y['unmatched'] or x['missing'] or y['missing']):flags.insert(0,'Сигнал снижения; сверить начисления')
            # Prefer recent data only when the observed cohort is sufficiently populated and has no missing calculation.
            eligible=lambda m:m['q']>=5 and not m['unmatched'] and not m['missing']
            base=x if eligible(x) else allm
            base_name='14д' if base is x else 'Весь период (14д неполны / мало шт.)'
            if not base['q'] or base['unmatched']:base_name='Нет полной базы для цены'
            def change(key):return x[key]-y[key] if x['q'] and y['q'] else None
            m=meta.get(sku,());v=[sku,names.get(sku,sku),model,x['q'],x['profit'],y['profit'],delta,allm['profit'],prices.get(sku),prices.get(sku),None,None,None,'; '.join(flags),y['q'],z['q'],z['profit'],x['missing'],y['missing'],x['unmatched'],y['unmatched'],base['cost'],base['profit'],base_name,change('buyer'),change('fees'),change('cost'),change('tax'),change('bonus'),x['cancel'],allm['q'],summaries[model,'recent'].get(sku,{}).get('returned',0),m[30] if m else None,m[33] if m else None,m[34] if m else None]
            ws.append([round(t,4) if isinstance(t,float) else t for t in v]);r=ws.max_row
            # Relative row references remain sortable in Excel; no payout-window division.
            if base['q'] and not base['unmatched'] and prices.get(sku):
                ws.cell(r,11,f'=IF(AND(ISNUMBER(J{r}),J{r}>0),(W{r}+V{r})*J{r}/I{r}-V{r},"")')
                ws.cell(r,12,f'=IF(W{r}>=100,I{r},IF(W{r}+V{r}>0,ROUNDUP(I{r}*(V{r}+100)/(W{r}+V{r})/10,0)*10,""))')
                ws.cell(r,13,f'=IF(ISNUMBER(L{r}),MAX(0,L{r}-I{r}),"")')
            if delta is not None:ws.cell(r,7,f'=E{r}-F{r}')
            for cell in ws[r]:
                cell.font=Font(name='Arial',size=10,bold=model=='Σ',color='333333');cell.alignment=Alignment(vertical='center',wrap_text=cell.column in (2,14,24));cell.number_format='#,##0.00;[Red](#,##0.00);"—"'
                if model=='Σ':cell.fill=PatternFill('solid',fgColor='E8EFF4')
                elif len(labels)>1:cell.fill=PatternFill('solid',fgColor='F6F8FA')
            for col in [4,15,16,18,19,20,21,31,32]:ws.cell(r,col).number_format='0'
            ws.cell(r,10).fill=PatternFill('solid',fgColor='FFF2CC');ws.cell(r,10).font=Font(color='0000FF',size=10)
            if len(flags)>1 or x['q']<5 or y['q']<5 or x['missing'] or y['missing'] or x['unmatched'] or y['unmatched']:ws.cell(r,14).fill=PatternFill('solid',fgColor='FFF2CC')
            ws.row_dimensions[r].height=60
        if len(labels)>1:ws.row_dimensions.group(first+1,ws.max_row,hidden=False)
    last=ws.max_row
    for c in ws[5]:c.fill=PatternFill('solid',fgColor='23465A');c.font=Font(bold=True,color='FFFFFF',size=10);c.alignment=Alignment(wrap_text=True,vertical='center',horizontal='center')
    ws.row_dimensions[5].height=58
    widths={'A':17,'B':30,'C':9,'D':11,'E':17,'F':17,'G':16,'H':17,'N':47,'X':38}
    for i in range(1,36):ws.column_dimensions[get_column_letter(i)].width=widths.get(get_column_letter(i),18)
    ws.column_dimensions.group('I','M',hidden=True);ws.column_dimensions.group('O','AI',hidden=True)
    ws.sheet_properties.outlinePr.summaryBelow=False;ws.sheet_properties.outlinePr.summaryRight=False;ws.freeze_panes='D6';ws.auto_filter.ref=f'A5:AI{last}'
    style_profit_view(ws,last)
    notes=[
      'Сортировка по умолчанию: SKU → наименование → итог Σ, FBS, FBY. Фильтр заголовка позволяет сортировать и по названию. При ручной сортировке строки каналов могут разойтись; SKU и канал указаны в каждой строке.',
      'Окна 14д не пересекаются. Дата заказа одинакова по смыслу в обоих окнах; поздние начисления и более короткий срок наблюдения свежих заказов могут смещать сравнение. Ставка налога — 7%.',
      'Δ = последние 14д − предыдущие 14д. Отрицательная Δ — сигнал потери на штуку, но при неполных проводках это не доказанное ухудшение экономики. Неизвестные значения остаются пустыми, а не нулём.',
      'Порог 5 шт. — фильтр малой выборки, не статистическое доказательство. 7д — только ранний сигнал. Разница FBS/FBY также может зависеть от цены, скидки и состава заказов, а не только логистики.',
      'Разложение Δ справа: изменение поступлений − услуг − закупа − налога + сальдо баллов. Все компоненты рассчитаны по тем же доставленным заказам. Буст уже внутри услуг и баллов, второй раз не вычитается.',
      'База цены: 14д только при ≥5 доставленных шт., без пропущенных расчётов и без доставок без проводок; иначе весь период с явной подписью X. Жёлтая J — цена сценария; L — модельная цена для 100 ₽/шт. без возвратов.',
      'База за весь период и сценарии тоже предварительные: наличие проводки не доказывает окончательность всех начислений. Возвраты, невыкупы и накладные расходы не входят в D–H; известный вклад невыкупов показан справа.',
      'Σ = сумма вкладов / сумма штук, а не среднее двух каналов. Не суммируйте Σ и строки FBS/FBY. Сверка выплат — деньги по датам проводок; закрытый сентябрь показан отдельно.',
    ]
    for i,note in enumerate(notes,last+3):ws.merge_cells(start_row=i,start_column=1,end_row=i,end_column=14);ws.cell(i,1,note);ws.cell(i,1).alignment=Alignment(wrap_text=True,vertical='center');ws.row_dimensions[i].height=38
    from .profit_drivers import add_profit_drivers
    add_profit_drivers(wb,orders,records,date_from,date_to)
    return ws


def style_profit_view(ws,last=None):
    """Profit bands apply to monetary profit levels, never to their change."""
    from openpyxl.formatting.rule import FormulaRule
    if last is None:
        from openpyxl.utils.cell import range_boundaries
        last=range_boundaries(ws.auto_filter.ref)[3]
    ws.conditional_formatting._cf_rules.clear()
    white=PatternFill('solid',fgColor='FFFFFF')
    active=PatternFill('solid',fgColor='EAF2F8')
    red=PatternFill('solid',fgColor='FBE3E2')
    yellow=PatternFill('solid',fgColor='FFF2CC')
    green=PatternFill('solid',fgColor='DFF0E4')
    # Highest priority: no sales makes the entire data row white, even with older profit values.
    ws.conditional_formatting.add(f'A6:AI{last}',FormulaRule(formula=['$D6=0'],fill=white,stopIfTrue=True))
    for row in ws.iter_rows(min_row=6,max_row=last):
        for cell in row:
            if row[3].value==0:cell.fill=white
            elif cell.column!=10 and cell.column!=14:cell.fill=active
    for col in ['E','F','H','K','Q','W']:
        target=f'{col}6:{col}{last}'
        for expression,fill,color in [
          (f'AND($D6>0,ISNUMBER({col}6),{col}6<50)',red,'A61C1C'),
          (f'AND($D6>0,ISNUMBER({col}6),{col}6>=50,{col}6<=100)',yellow,'7D5A00'),
          (f'AND($D6>0,ISNUMBER({col}6),{col}6>100)',green,'1B5E20')]:
            ws.conditional_formatting.add(target,FormulaRule(formula=[expression],fill=fill,font=Font(color=color),stopIfTrue=True))
    # Change is directional, so a +20 delta must not look like low profit.
    for expr,fill,color in [('AND($D6>0,ISNUMBER(G6),G6<0)',red,'A61C1C'),('AND($D6>0,ISNUMBER(G6),G6>=0)',green,'1B5E20')]:
        ws.conditional_formatting.add(f'G6:G{last}',FormulaRule(formula=[expr],fill=fill,font=Font(color=color),stopIfTrue=True))
    ws.conditional_formatting.add(f'A6:AI{last}',FormulaRule(formula=['$D6>0'],fill=active))
