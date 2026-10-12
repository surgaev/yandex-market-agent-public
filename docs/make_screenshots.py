"""Скриншоты листов отчёта для README — на ВЫДУМАННЫХ данных.

Рисует HTML-макеты листов «Сверка выплат» и «Прибыль по месяцам» с теми же
цветами и правилами подсветки, что в настоящем отчёте, и снимает их
безголовым Chrome в PNG. Данные магазина сюда не попадают: артикулы, названия
и цифры сгенерированы ниже с фиксированным seed.

Запуск (macOS): python3 docs/make_screenshots.py
"""
import os
import random
import re
import subprocess

HERE = os.path.dirname(os.path.abspath(__file__))
CHROME = "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome"
rnd = random.Random(7)

GREEN, YELLOW, RED = ("#C6EFCE", "#1B5E20"), ("#FFF9C4", "#7D6608"), ("#FFC7CE", "#B71C1C")


def tier(v):
    return GREEN if v > 100 else YELLOW if v >= 50 else RED


def fmt(v, dec=0):
    if v is None:
        return ""
    s = f"{v:,.{dec}f}".replace(",", " ")
    return s.replace("-", "−")


CSS = """
body{margin:0;background:#fff;font-family:Calibri,'Segoe UI',Arial,sans-serif;font-size:12px;color:#000}
.wrap{padding:10px 12px 0 12px;display:inline-block}
.note{color:#555;font-size:11px;max-width:1100px;white-space:normal;margin:0 0 6px 2px;line-height:1.35}
table{border-collapse:collapse}
td,th{border:1px solid #c9c9c9;padding:3px 6px;white-space:nowrap;height:18px}
th{color:#fff;font-weight:700;text-align:center;white-space:pre-line;line-height:1.15;font-size:11px}
td.l{text-align:left} td{text-align:right}
td.c{text-align:center}
b{font-weight:700}
.tabs{margin-top:8px;border-top:1px solid #bbb;padding:4px 0 6px 2px;font-size:11px;color:#333}
.tabs span{display:inline-block;padding:3px 10px;margin-right:2px;border:1px solid #c9c9c9;border-top:none;background:#f3f3f3}
.tabs span.on{background:#fff;font-weight:700;border-bottom:2px solid #217346;color:#217346}
"""

TABS = ["Дашборд", "Скидки FEE по заказам", "Сверка выплат", "Прибыль по месяцам",
        "Юнит-экономика", "Товары", "FBY остатки по дням", "Показы 14д", "Покупки 14д", "…"]


def tabs(active):
    return '<div class="tabs">' + "".join(
        f'<span class="{"on" if t == active else ""}">{t}</span>' for t in TABS) + "</div>"


# ── выдуманный каталог: (название, артикул, цена, себестоимость, продаж за 90д) ──
CATALOG = [
    ("Чехол для телефона силиконовый", "CASE-SIL-01", 890, 140, 118),
    ("Ремешок для часов кожаный", "STRAP-LTH-BK", 1490, 310, 64),
    ("Держатель для телефона в авто", "HOLD-CAR-02", 1290, 260, 92),
    ("Кабель USB-C плетёный 2 м", "CBL-USBC-2M", 590, 85, 155),
    ("Перчатки сенсорные, L", "GLV-TCH-L", 2190, 420, 47),
    ("Наушники-вкладыши", "EAR-BUD-W", 2990, 1150, 38),
    ("Подставка для ноутбука", "STAND-LP-AL", 2490, 980, 21),
]


PPU90 = {}   # sku → прибыль на штуку за 90 дней (из «Сверки»), чтобы листы не противоречили друг другу


def svverka_html():
    heads = [("Товар", "#C0392B"), ("SKU", "#C0392B"), ("Модель", "#C0392B"),
             ("Чистая\nприбыль\nФАКТ ₽/шт", "#1F3A5F"), ("⭐ ВЗВЕШ.\n₽/шт", "#1F3A5F"),
             ("Продано\nшт.", "#C0392B"), ("Возвр.\nшт.", "#C0392B"), ("Платежи\nпокуп. ₽", "#C0392B"),
             ("Удержано\nуслуги ₽", "#C0392B"), ("Скидка\nакций ₽", "#C0392B"), ("Баллы-\nкомпенс. ₽", "#C0392B"),
             ("К перечисл.\nФАКТ ₽", "#C0392B"), ("Себест. ₽", "#C0392B"), ("Налог\n7% ₽", "#C0392B"),
             ("Чистая\nприбыль\nФАКТ ₽", "#C0392B"), ("Маржа %", "#C0392B")]
    rows = []
    tot = dict(sold=0, ret=0, pay=0, svc=0, disc=0, pts=0, out=0, cost=0, tax=0, prof=0)
    for gi, (name, sku, price, cost_u, sold90) in enumerate(CATALOG):
        band = "#F2F2F2" if gi % 2 == 0 else "#FFFFFF"
        models = [("FBS", 0.65), ("FBY", 0.35)] if sold90 > 40 else [("FBS", 1.0)]
        agg = dict(sold=0, ret=0, pay=0, svc=0, disc=0, pts=0, out=0, cost=0, tax=0, prof=0)
        for model, share in models:
            sold = max(1, round(sold90 / 3 * share * rnd.uniform(0.8, 1.2)))
            ret = rnd.choice([0, 0, 0, 1]) if sold > 4 else 0
            pay = sold * price * rnd.uniform(0.42, 0.55)
            disc = -sold * price * rnd.uniform(0.45, 0.52)
            pts = -disc * rnd.uniform(0.97, 1.06)
            svc = -sold * price * rnd.uniform(0.10, 0.17) - sold * rnd.uniform(60, 120)
            retm = -ret * price * 0.5
            out = pay + disc + pts + svc + retm
            cost = (sold - ret) * cost_u
            tax = 0.07 * (pay + retm)
            prof = out - cost - tax
            ppu = prof / (sold - ret)
            vz = ppu * rnd.uniform(0.85, 1.1)
            for k, v in zip(agg, (sold, ret, pay, svc, disc, pts, out, cost, tax, prof)):
                agg[k] += v
            rows.append((band, False, name, sku, model, ppu, vz, sold, ret, pay, svc, disc, pts, out, cost, tax, prof,
                         prof / pay * 100 if pay else 0))
        nq = agg["sold"] - agg["ret"]
        ppu90 = agg["prof"] / nq
        PPU90[sku] = ppu90   # тот же уровень прибыли на штуку — для «Прибыли по месяцам»
        rows.append(("#FDEBD0", True, "", sku, "Σ 90д", ppu90, ppu90 * rnd.uniform(0.92, 1.05), agg["sold"] * 3,
                     agg["ret"] * 3, agg["pay"] * 3, None, agg["disc"] * 3, agg["pts"] * 3, agg["out"] * 3,
                     agg["cost"] * 3, agg["tax"] * 3, agg["prof"] * 3, agg["prof"] / agg["pay"] * 100))
        for k in tot:
            tot[k] += agg[k]
    nq = tot["sold"] - tot["ret"]
    rows.append(("#F2F2F2", True, "ИТОГО", "", "", tot["prof"] / nq, tot["prof"] / nq * 0.9, tot["sold"], tot["ret"],
                 tot["pay"], tot["svc"], tot["disc"], tot["pts"], tot["out"], tot["cost"], tot["tax"], tot["prof"],
                 tot["prof"] / tot["pay"] * 100))

    h = ['<tr>' + "".join(f'<th style="background:{c}">{t}</th>' for t, c in heads) + '</tr>']
    for band, bold, name, sku, model, ppu, vz, sold, ret, pay, svc, disc, pts, out, cost, tax, prof, mrg in rows:
        bw = "font-weight:700;" if bold else ""
        def cell(v, dec=0, cls=""):
            return f'<td class="{cls}" style="background:{band};{bw}">{fmt(v, dec)}</td>'
        def tiered(v):
            bg, fg = tier(v)
            return f'<td style="background:{bg};color:{fg};font-weight:700">{fmt(v)}</td>'
        pc = "#1B5E20" if prof > 0 else "#B71C1C"
        h.append("<tr>" + f'<td class="l" style="background:{band};{bw}">{name}</td>'
                 + f'<td class="l" style="background:{band};{bw}">{sku}</td>'
                 + f'<td class="c" style="background:{band};{bw}">{model}</td>'
                 + tiered(ppu) + tiered(vz)
                 + cell(sold) + cell(ret) + cell(pay) + cell(svc) + cell(disc) + cell(pts) + cell(out)
                 + cell(cost) + cell(tax)
                 + f'<td style="background:{band};color:{pc};font-weight:700">{fmt(prof)}</td>'
                 + f'<td style="background:{band};color:{pc};font-weight:700">{fmt(mrg, 1)}</td></tr>')
    note = ("РЕАЛЬНЫЕ выплаты от Яндекса по артикулам (отчёт «По платежам») за 30 дней. Чистая прибыль ФАКТ = "
            "к перечислению − себестоимость (на выкупленные штуки) − налог 7%. Оранжевая строка «Σ 90д» — "
            "устойчивая квартальная прибыль. ⭐ ВЗВЕШ. ₽/шт — главная цифра: 0–30 и 31–90 дней с весом 2:1. "
            "Подсветка: >100 ₽ зелёный, 50–100 ₽ жёлтый, <50 ₽ красный. <i>Данные выдуманные.</i>")
    return (f"<html><head><meta charset='utf-8'><style>{CSS}</style></head><body><div class='wrap'>"
            f"<div class='note'>{note}</div><table>{''.join(h)}</table>{tabs('Сверка выплат')}</div></body></html>")


def months_html():
    months = ["2026-05", "2026-06", "2026-07", "2026-08", "2026-09", "2026-10"]
    TQ, TR, TP = "#E8F1F8", "#FDF3E2", "#EDE7F6"
    data = []
    for name, sku, price, cost_u, sold90 in CATALOG:
        base = sold90 / 3
        ppu0 = PPU90[sku]
        cells = []
        for i, m in enumerate(months):
            if rnd.random() < 0.08:
                cells.append((None, None))
                continue
            q = max(0, round(base * rnd.uniform(0.5, 1.4) * (0.45 if i == 5 else 1)))
            # к текущему месяцу ₽/шт выходит на уровень «Сверки»; раньше — с разбросом
            p = q * (ppu0 + (5 - i) * rnd.uniform(-12, 12) + rnd.uniform(-25, 25)) if q else None
            cells.append((q if q else None, p))
        data.append((name, sku, cells))
    data.sort(key=lambda d: -sum(p or 0 for _, p in d[2]))
    grand = sum(sum(p or 0 for _, p in d[2]) for d in data)

    def pf(v, bold=False):
        c = "#1B5E20" if v > 100 else "#7D6608" if v >= 50 else "#B71C1C"
        return f"color:{c};" + ("font-weight:700;" if bold else "")

    h = ['<tr><th rowspan=2 style="background:#8E44AD">Товар</th><th rowspan=2 style="background:#8E44AD">SKU</th>'
         + "".join(f'<th colspan=3 style="background:#8E44AD">{m}</th>' for m in months)
         + "".join(f'<th rowspan=2 style="background:#8E44AD">{t}</th>' for t in ("Итого ₽", "Итого\nшт.", "Итого\n₽/шт", "Доля %"))
         + "</tr>",
         "<tr>" + "".join(f'<th style="background:{c};color:#333">{t}</th>' for _ in months
                          for t, c in (("шт.\nчистые", TQ), ("прибыль ₽", TR), ("₽/шт", TP))) + "</tr>"]
    mt = [[0, 0] for _ in months]
    for name, sku, cells in data:
        tot_p = sum(p or 0 for _, p in cells)
        tot_q = sum(q or 0 for q, p in cells if p is not None)
        share = tot_p / grand * 100
        row = f'<td class="l">{name}</td><td class="l">{sku}</td>'
        for i, (q, p) in enumerate(cells):
            row += f'<td style="background:{TQ}">{fmt(q)}</td>'
            if p is None:
                row += f'<td style="background:{TR}"></td><td style="background:{TP}"></td>'
                continue
            mt[i][0] += q or 0
            mt[i][1] += p
            row += f'<td style="background:{TR};color:{"#1B5E20" if p > 0 else "#B71C1C"}">{fmt(p)}</td>'
            ppu = p / q if q else None
            row += f'<td style="background:{TP};{pf(ppu) if ppu is not None else ""}">{fmt(ppu)}</td>'
        sbg = "#C6EFCE" if share >= 10 else "#FFF9C4" if share >= 5 else "#fff"
        row += (f'<td style="font-weight:700;color:{"#1B5E20" if tot_p > 0 else "#B71C1C"}">{fmt(tot_p)}</td>'
                f'<td style="background:{TQ}">{fmt(tot_q)}</td>'
                f'<td style="background:{TP};{pf(tot_p / tot_q, True)}">{fmt(tot_p / tot_q)}</td>'
                f'<td style="background:{sbg}">{fmt(share, 1)}</td>')
        h.append(f"<tr>{row}</tr>")
    g = "background:#F2F2F2;font-weight:700"
    row = f'<td class="l" style="{g}">ИТОГО</td><td style="{g}"></td>'
    for q, p in mt:
        row += (f'<td style="{g}">{fmt(q)}</td><td style="{g}">{fmt(p)}</td>'
                f'<td style="{g};{pf(p / q, True) if q else ""}">{fmt(p / q) if q else ""}</td>')
    gq = sum(q for q, _ in mt)
    row += (f'<td style="{g}">{fmt(grand)}</td><td style="{g}">{fmt(gq)}</td>'
            f'<td style="{g};{pf(grand / gq, True)}">{fmt(grand / gq)}</td><td style="{g}"></td>')
    h.append(f"<tr>{row}</tr>")
    mom = '<td class="l" style="font-weight:700">MoM %</td><td></td>'
    for i in range(len(months)):
        if i == 0:
            mom += "<td></td><td></td><td></td>"
            continue
        prev, cur = mt[i - 1][1], mt[i][1]
        pct = (cur - prev) / prev * 100 if prev else 0
        mom += (f'<td></td><td style="font-weight:700;color:{"#1B5E20" if pct >= 0 else "#B71C1C"}">'
                f'{"+" if pct >= 0 else "−"}{abs(pct):.1f}%</td><td></td>')
    h.append(f"<tr>{mom}<td></td><td></td><td></td><td></td></tr>")
    note = ("Помесячная чистая прибыль ФАКТ по артикулам: на каждый месяц — чистые штуки (продано − возвраты), "
            "прибыль ₽ и ₽/шт. «Итого ₽/шт» = прибыль за все месяцы ÷ штуки за все месяцы. Закрытые месяцы "
            "кэшируются, последний — текущий, обновляется каждый прогон. <i>Данные выдуманные.</i>")
    return (f"<html><head><meta charset='utf-8'><style>{CSS}</style></head><body><div class='wrap'>"
            f"<div class='note'>{note}</div><table>{''.join(h)}</table>{tabs('Прибыль по месяцам')}</div></body></html>")


MEASURE = ("<script>window.onload=()=>{const w=document.querySelector('.wrap');"
           "document.body.setAttribute('data-size',Math.ceil(w.scrollWidth+24)+'x'+Math.ceil(w.scrollHeight+4));}"
           "</script>")


def shoot(html, name):
    """Два прохода: Chrome сначала меряет страницу (--dump-dom), потом снимает
    ровно по её размеру — без обрезки и пустых полей."""
    path = os.path.join(HERE, name + ".html")
    open(path, "w", encoding="utf-8").write(html.replace("</body>", MEASURE + "</body>"))
    common = [CHROME, "--headless=new", "--disable-gpu", "--hide-scrollbars"]
    dom = subprocess.run(common + ["--window-size=4000,3000", "--dump-dom", "file://" + path],
                         check=True, capture_output=True, text=True).stdout
    w, h = re.search(r'data-size="(\d+)x(\d+)"', dom).groups()
    png = os.path.join(HERE, name + ".png")
    subprocess.run(common + ["--force-device-scale-factor=2", f"--window-size={w},{h}",
                             f"--screenshot={png}", "file://" + path],
                   check=True, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    os.remove(path)
    print(f"→ {png} ({w}×{h})")


if __name__ == "__main__":
    shoot(svverka_html(), "screenshot-sverka")
    shoot(months_html(), "screenshot-months")
