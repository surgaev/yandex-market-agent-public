"""Explain delivered-unit economics and separately disclose return/cancellation costs."""
from .config import DATA_DIR
from collections import defaultdict,Counter
from datetime import date,timedelta
from pathlib import Path
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils.cell import range_boundaries
from .sku_diagnostics import summarize_skus

GROUPS={
 'Комиссии':['Размещение товаров на витрине, ₽','Приём платежа покупателя, ₽','Перевод платежа покупателя, ₽','Вознаграждение за продажу товара, ₽'],
 'Доставка':['Доставка покупателю, ₽','Доставка (средняя миля), ₽','Экспресс-доставка покупателю, ₽','Доставка из-за рубежа, ₽','Организация забора заказов, ₽'],
 'Буст':['Буст продаж, ₽'],
 'Обработка/хранение':['Складская обработка, ₽','Обработка заказов в СЦ или ПВЗ, ₽','Хранение невыкупов и возвратов, ₽','Обработка заказов на складе, ₽'],
 'Другие услуги':['Программа лояльности и отзывы, ₽','Рассрочка, ₽'],
}

def read_fee_breakdown(path):
    if not path.exists():return {}
    wb=openpyxl.load_workbook(path,read_only=True,data_only=True)
    try:
        ws=wb['Услуги и маржа по заказам'];ws.reset_dimensions();it=iter(ws.values)
        for header in it:
            if 'Номер заказа' in header and 'Все услуги Маркета за заказы, ₽' in header:break
        else:raise ValueError('Нет шапки детализации услуг')
        index={label:i for i,label in enumerate(header) if label is not None}
        required=['Модели работы','Номер заказа','Все услуги Маркета за заказы, ₽']+[h for labels in GROUPS.values() for h in labels]
        missing=[h for h in required if h not in index]
        if missing:raise ValueError('Неизвестная схема услуг: '+', '.join(missing))
        out={}
        for row in it:
            if len(row)<=index['Номер заказа'] or row[index['Номер заказа']] is None:continue
            model=str(row[index['Модели работы']]);oid=row[index['Номер заказа']];oid=str(int(oid)) if isinstance(oid,(float,int)) else str(oid)
            get=lambda h:float(row[index[h]] or 0)
            parts={group:sum(get(h) for h in labels)for group,labels in GROUPS.items()}
            total=get('Все услуги Маркета за заказы, ₽');parts['Другие услуги']+=total-sum(parts.values())
            key=(model,oid)
            if key in out:raise ValueError('Повтор заказа в детализации')
            out[key]=(total,parts)
        return out
    finally:wb.close()


def explain(orders,aggregate,breakdown,counts):
    q=aggregate.get('qty',0);parts=defaultdict(float);covered=0;eligible=0;known_fees=0
    for r in orders:
        if r.get('status')!='Доставлен покупателю' or r.get('after_reserve') is None:continue
        eligible+=1;key=(r['model'],str(r['order']));item=breakdown.get(key)
        if item is None:continue
        total,detail=item
        if abs(total)<1e-9:
            if any(abs(v)>1e-9 for v in detail.values()) and counts[key]>1:continue
            weight=1 if counts[key]==1 else 0
        else:weight=float(r['fees'])/total
        for label,value in detail.items():parts[label]+=value*weight
        covered+=1;known_fees+=r['fees'] or 0
    def rub(value):return f'{value:.1f}'.replace('.',',')
    if q:
        cost=aggregate.get('cost',0)/q;tax=(aggregate.get('buyer',0)-aggregate.get('fees',0)-aggregate.get('cost',0)-aggregate.get('delivered_cash',0))/q
        lines=[f"Поступления {rub(aggregate.get('buyer',0)/q)} ₽/шт.; закуп {rub(cost)}; налог {rub(tax)}."]
        if covered:
            ranked=sorted(parts.items(),key=lambda kv:-kv[1]);text='; '.join(f'{name} {rub(value/q)}' for name,value in ranked if abs(value)>=.005)
            lines.append((text or 'Услуги 0,0')+' ₽/шт.'+(' (разбивка неполная)' if covered!=eligible else ''))
        else:lines.append('Услуги '+rub(aggregate.get('fees',0)/q)+' ₽/шт.; разбивка недоступна.')
        if covered and abs(known_fees-sum(parts.values()))>.011:raise ValueError('Детализация не сходится с услугами')
        bonus=(aggregate.get('delivered_adjusted',0)-aggregate.get('delivered_cash',0))/q
        if abs(bonus)>=.005:lines.append('Сальдо баллов '+('+' if bonus>0 else '')+rub(bonus)+' ₽/шт.; может уточняться.')
    else:lines=['Нет рассчитанных доставок; причину ₽/шт. определить нельзя.']
    cancelled=aggregate.get('cancelled',0)
    if cancelled:
        amount=aggregate.get('cancel_adjusted',0)
        lines.append(f"Невыкупы/отмены: {int(cancelled)} поз., вклад {rub(amount)} ₽ итого"+(f' ({rub(amount/q)} ₽/доставл.шт.)' if q else '')+' — отдельно.')
    returned=aggregate.get('returned',0)
    if returned:lines.append(f"Возвраты: {int(returned)} поз.; известные услуги {rub(aggregate.get('return_fees_known',0))} ₽ итого — в ₽/шт. не включены.")
    if aggregate.get('missing',0):lines.append(f"Не рассчитано {int(aggregate['missing'])} поз.; вывод предварительный.")
    return '\n'.join(lines)


def add_profit_drivers(wb,orders,records,date_from,date_to,raw_path=None):
    ws=wb['Юнит-экономика'];last=range_boundaries(ws.auto_filter.ref)[3]
    start=date.fromisoformat(str(date_from));end=date.fromisoformat(str(date_to));recent=max(start,end-timedelta(days=13))
    path=Path(raw_path) if raw_path else Path(DATA_DIR)/'finance_cache_v2'/f'orders_{start}_{end}.xlsx'
    breakdown=read_fee_breakdown(path);counts=Counter((r['model'],str(r['order'])) for r in orders)
    computed={}
    for model in ['Σ','FBS','FBY']:
        subset=[r for r in orders if (model=='Σ' or r['model']==model) and str(start)<=r['created']<=str(end)]
        for label,lo in [('all',start),('recent',recent)]:
            cohort=[r for r in subset if str(lo)<=r['created']]
            agg=summarize_skus(cohort,records,start,end);by_sku=defaultdict(list)
            for r in cohort:by_sku[r['sku']].append(r)
            for sku,a in agg.items():computed[(model,label,sku)]=explain(by_sku[sku],a,breakdown,counts)
    for col,label in [(36,'Что съедает прибыль · 14 дней'),(37,'Что съедает прибыль · весь период')]:
        c=ws.cell(5,col,label);c.fill=PatternFill('solid',fgColor='23465A');c.font=Font(bold=True,color='FFFFFF',size=10);c.alignment=Alignment(wrap_text=True,vertical='center',horizontal='center')
        ws.column_dimensions[c.column_letter].width=54;ws.column_dimensions[c.column_letter].hidden=False
        for ri in range(6,last+1):
            sku=ws.cell(ri,1).value;model=ws.cell(ri,3).value
            value=computed.get((model,'recent' if col==36 else 'all',sku),'Нет заказов в этом окне.')
            c=ws.cell(ri,col,value);c.font=Font(name='Arial',size=10,color='333333');c.alignment=Alignment(wrap_text=True,vertical='center')
            c.fill=PatternFill('solid',fgColor='FFFFFF' if not ws.cell(ri,4).value else 'EAF2F8')
            # Keep every cost and missing-data warning readable.
            lines=sum(max(1,(len(line)+53)//54) for line in value.split('\n'))
            ws.row_dimensions[ri].height=max(ws.row_dimensions[ri].height or 60,lines*13+10)
    ws.auto_filter.ref=f'A5:AK{last}'
    rng=f'AJ6:AK{last}'
    ws.conditional_formatting.add(rng,FormulaRule(formula=['$D6=0'],fill=PatternFill('solid',fgColor='FFFFFF'),stopIfTrue=True))
    ws.conditional_formatting.add(rng,FormulaRule(formula=['$D6>0'],fill=PatternFill('solid',fgColor='EAF2F8')))
    return computed
