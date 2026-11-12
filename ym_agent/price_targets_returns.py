"""Price scenarios including observed cancellation and return losses, not a guarantee."""
from .config import DATA_DIR
import os
from pathlib import Path
from datetime import date,timedelta
from collections import Counter
import openpyxl
from .sku_diagnostics import summarize_skus
from .profit_drivers import read_fee_breakdown

def build_rows(wb,orders,records,date_from,date_to,meta=None,prices=None,raw_path=None,bids=None):
    end=date.fromisoformat(str(date_to))-timedelta(days=14)
    sku_prefix = os.environ.get('YM_PRICE_SKU_PREFIX', '')
    cohort=[r for r in orders if str(date_from)<=r['created']<=str(end)]
    agg=summarize_skus(cohort,records,date_from,date_to)
    prices=dict(prices or {});names={};skus=set()
    for name in ['Товары','Индекс цен']:
        if name not in wb.sheetnames:continue
        for r in wb[name].values:
            if len(r)>1 and str(r[1]).startswith(sku_prefix):
                sku=r[1];skus.add(sku);names[sku]=r[0]
                if name=='Индекс цен' and len(r)>2 and isinstance(r[2],(float,int)):prices[sku]=r[2]
    skus.update(r['sku'] for r in orders if r['sku'].startswith(sku_prefix))
    for r in records:
        if r['sku'] in skus and r['source']=='Платёж покупателя' and r.get('service'):
            if len(r['service'])>len(names.get(r['sku'],'')):names[r['sku']]=r['service']
    bids=dict(bids or {})
    for sku,m in (meta or {}).items():
        if len(m)>30 and isinstance(m[30],(int,float)):bids.setdefault(sku,m[30])
    path=Path(raw_path) if raw_path else Path(DATA_DIR)/'finance_cache_v2'/f'orders_{date_from}_{date_to}.xlsx'
    detail=read_fee_breakdown(path) if path.exists() else {};gross={}
    if path.exists():
        raw=openpyxl.load_workbook(path,read_only=True,data_only=True);ws=raw['Транзакции по заказам и товарам'];ws.reset_dimensions();it=iter(ws.values)
        for h in it:
            if 'Ваш SKU' in h and 'Цена продажи (за шт.), ₽' in h:break
        idx={k:h.index(k) for k in ['Ваш SKU','Номер заказа','Модели работы','Цена продажи (за шт.), ₽']}
        for rr in it:
            oid=rr[idx['Номер заказа']]
            if oid is None:continue
            oid=str(int(oid)) if isinstance(oid,(int,float)) else str(oid)
            gross[(str(rr[idx['Модели работы']]),oid,str(rr[idx['Ваш SKU']]))]=rr[idx['Цена продажи (за шт.), ₽']]
        raw.close()
    counts=Counter((r['model'],str(r['order'])) for r in orders);rows=[]
    for sku in sorted(skus):
        a=agg.get(sku,{});q=a.get('qty',0);g=0;boost=0;covered=0;boostcovered=0
        for r in cohort:
            if r['sku']!=sku or r['status']!='Доставлен покупателю' or r['after_reserve'] is None:continue
            key=(r['model'],str(r['order']));price=gross.get((*key,sku))
            if isinstance(price,(int,float)) and price>0:g+=price*r['qty'];covered+=r['qty']
            d=detail.get(key)
            if d and (d[0]!=0 or counts[key]==1):
                weight=r['fees']/d[0] if d[0] else 1;boost+=d[1]['Буст']*weight;boostcovered+=r['qty']
        ok=q>0 and abs(covered-q)<1e-6 and not a.get('delivered_unmatched')
        note=[]
        if not ok:note.append('Нет полной базы для цены')
        if q<5:note.append('Менее 5 доставок: слабая оценка')
        if a.get('missing'):note.append('Незавершённые/нерассчитанные позиции: '+str(int(a['missing'])))
        note.append('Допрасходы не заданы; скидки/возвраты могут измениться')
        pu=lambda k:a.get(k,0)/q if q else None
        # Input columns A:F; outputs G:N; calculation inputs O:W.
        rows.append([sku,names.get(sku,sku),prices.get(sku),q,bids.get(sku),None,None,None,None,None,None,None,None,'; '.join(note),g/q if ok else None,pu('cost'),pu('delivered_adjusted'),max(0,-a.get('cancel_adjusted',0))/q if q else None,a.get('return_fees_known',0)/q if q else None,boost/q if q and abs(boostcovered-q)<1e-6 else None,str(date_from)+'–'+str(end),a.get('returned',0),a.get('cancelled',0)])
    return rows

def formulas(r):
    # Historical retention incl tax/fees/bonus, fixed cost and known loss allowance.
    k=f'((Q{r}+P{r})/O{r})'
    fixed=f'(P{r}+R{r}+S{r}+IF(ISNUMBER(F{r}),F{r},0))'
    stress=f'({k}-MAX(0,E{r}/100-T{r}/O{r}))'
    valid=f'AND(ISNUMBER(C{r}),C{r}>0,ISNUMBER(O{r}),O{r}>0,ISNUMBER(Q{r}))'
    return [
      f'=IF({valid},{k}*C{r}-{fixed},"н.д.")',
      f'=IF({valid},IF({k}>0,MAX(C{r},ROUNDUP((100+{fixed})/{k}/10,0)*10),"нет решения"),"н.д.")',
      f'=IF(ISNUMBER(H{r}),H{r}-C{r},"н.д.")',
      f'=IF(AND({valid},ISNUMBER(E{r}),ISNUMBER(T{r})),IF({stress}>0,MAX(C{r},ROUNDUP((100+{fixed})/{stress}/10,0)*10),"нет решения"),"н.д.")',
      f'=IF(ISNUMBER(J{r}),J{r}-C{r},"н.д.")',
      f'=IF(ISNUMBER(H{r}),{k}*H{r}-{fixed},"н.д.")',
      f'=IF(ISNUMBER(J{r}),{stress}*J{r}-{fixed},"н.д.")']
HEADERS=['SKU','Цвет / товар','Текущая цена, ₽','Доставок базы, шт.','Буст сейчас, %','Допрасходы, ₽/выкуп','Прибыль при текущей цене, модель','Цена для 100 ₽: прежние удержания','Поднять на, ₽','Цена для 100 ₽: буст полностью деньгами','Поднять на, ₽ (буст)','Проверка прибыли: H','Проверка прибыли: J','Ограничения','Средняя цена продаж базы, ₽','Закуп, ₽/шт.','Вклад доставок, ₽/шт.','Потери невыкупов, ₽/шт.','Услуги возвратов, ₽/шт.','Буст в услугах, ₽/шт.','Период заказов базы','Возвратов, поз.','Невыкупов, поз.']
NOTES=[
'Цель 100 ₽ на выкупленную штуку. Это сценарные цены при сохранении структуры скидок, услуг, налогов и доли возвратов; не гарантия чистой прибыли.',
'H: прежняя доля поступлений после услуг, налога и баллов; закуп и средние потери возвратов/невыкупов фиксированы. Использована историческая цена продажи из отчёта, а не сегодняшняя цена.',
'J: дополнительно резерв на полную денежную оплату текущей ставки буста для каждой продажи сверх буста, уже включённого в услуги. Это стресс-сценарий без гарантированной скидки баллами, не ожидаемое списание.',
'Закуп, налог 7% и учтённые услуги включены. Неизвестные расходы, будущие удержания и налог на возможную премию не оценены. Возвраты предполагаются годными к перепродаже.',
'Жёлтая F — дополнительные затраты на выкупленную штуку (упаковка, сборка, накладные). Пусто = не заданы; модель временно использует 0, поэтому нельзя считать результат прибылью после абсолютно всех расходов.',
'Цена не снижается ниже текущей, округляется вверх до 10 ₽. Цена продажи после скидки продавца может отличаться от прайс-листа; если акция уменьшает её, корректируйте цену/участие отдельно.',
'Цены в кабинете не меняются. J — стресс-проверка, не рекомендация: экстремальная цена означает необходимость пересмотреть буст/расходы. Н.д. = нет полной базы. Достигнуть 100 ₽ не гарантируется.'
]

def add_price_targets_returns(wb,orders,records,date_from,date_to,meta=None,prices=None,raw_path=None,bids=None):
    from openpyxl.styles import Font,PatternFill,Alignment
    from openpyxl.utils import get_column_letter
    name='Цены по товарам'
    if name in wb.sheetnames:del wb[name]
    ws=wb.create_sheet(name,wb.sheetnames.index('Юнит-экономика')+1)
    ws.append(HEADERS);header=ws.max_row
    for row in build_rows(wb,orders,records,date_from,date_to,meta,prices,raw_path,bids):
        ws.append(row);r=ws.max_row
        for j,f in enumerate(formulas(r),7):ws.cell(r,j,f)
        for c in ws[r]:c.font=Font(name='Arial',size=10);c.alignment=Alignment(wrap_text=c.column in [2,14],vertical='center');c.number_format='#,##0.0;[Red](#,##0.0);"—"'
        ws.cell(r,6).fill=PatternFill('solid',fgColor='FFF1CC');ws.cell(r,6).font=Font(color='0000FF');ws.row_dimensions[r].height=80
    for c in ws[header]:c.fill=PatternFill('solid',fgColor='23465A');c.font=Font(bold=True,color='FFFFFF');c.alignment=Alignment(wrap_text=True)
    ws.row_dimensions[header].height=65
    for i in range(1,24):ws.column_dimensions[get_column_letter(i)].width=20
    ws.column_dimensions['B'].width=38;ws.column_dimensions['N'].width=60;ws.freeze_panes='C2';ws.auto_filter.ref=f'A1:W{ws.max_row}';ws.sheet_view.showGridLines=False
    from openpyxl.formatting.rule import FormulaRule
    last=ws.max_row
    if last >= 2:
        for condition,fill,color in [('G2>100','DFF0E4','1B5E20'),('AND(G2>=50,G2<=100)','FFF2CC','7D5A00'),('G2<50','FBE3E2','A61C1C')]:
            ws.conditional_formatting.add(f'G2:G{last}',FormulaRule(formula=[f'AND(ISNUMBER(G2),{condition})'],fill=PatternFill('solid',fgColor=fill),font=Font(color=color)))
    for i,note in enumerate(['Цены по товарам с учётом известных потерь возвратов и невыкупов']+NOTES,last+2):
        ws.cell(i,1,note);ws.merge_cells(start_row=i,start_column=1,end_row=i,end_column=14)
        ws.row_dimensions[i].height=33;ws.cell(i,1).alignment=Alignment(wrap_text=True,vertical='center')
    return ws
