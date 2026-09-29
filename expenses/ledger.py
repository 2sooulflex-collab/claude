#!/usr/bin/env python3
"""Учет расходов 2Flex2Sooul: данные в CSV -> рабочая таблица xlsx для Google Таблиц.

Команды:
  add [файл.json]   добавить платежи (JSON-массив, как в промте; без файла — из stdin)
  batch "Партия" units=1780 boxes=130 per_box=14 liters=60 shipped=500
                    обновить количества на листе Партии
  check             проверить данные против справочников
  build             собрать out/Расходы 2Flex2Sooul.xlsx
  summary           посчитать Сводку в Python (для сверки с таблицей и ответов в чате)
  import файл.xlsx  забрать Транзакции и Партии из выгрузки Google Таблицы (правки руками, восстановление)
"""
import csv
import datetime as dt
import json
import sys
from pathlib import Path
from zoneinfo import ZoneInfo

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent
LISTS = json.loads((ROOT / "lists.json").read_text(encoding="utf-8"))
DATA = ROOT / "data"
TX_CSV = DATA / "transactions.csv"
BATCH_CSV = DATA / "batches.csv"
OUT = ROOT / "out" / "Расходы 2Flex2Sooul.xlsx"

# колонки листа Транзакции: ключ, заголовок, ключ JSON из промта
TX_FIELDS = [
    ("date", "Дата", "дата"),
    ("cat", "Категория расхода", "категория"),
    ("art", "Статья", "статья"),
    ("ptype", "Тип изделия", "тип_изделия"),
    ("season", "Сезон", "сезон"),
    ("supplier", "Поставщик/производство", "поставщик"),
    ("batch", "Артикул/партия", "партия"),
    ("alloc", "Тип отнесения", "тип_отнесения"),
    ("boxes", "Коробок, шт", "коробок"),
    ("amount", "Сумма", "сумма"),
    ("comment", "Комментарий", "комментарий"),
]
TX_COLS = [f[1] for f in TX_FIELDS]
C = {f[0]: i for i, f in enumerate(TX_FIELDS)}
JSON_KEYS = {f[2]: i for i, f in enumerate(TX_FIELDS)}
BATCH_COLS = ["Артикул/партия", "Тип изделия", "Сезон", "Поставщик/производство",
              "Единиц в партии", "Коробок, шт", "Единиц в коробке", "Литраж коробки",
              "Отгружено единиц"]
BATCH_KEYS = {"units": 4, "boxes": 5, "per_box": 6, "liters": 7, "shipped": 8}
OTHER = "Прочие расходы"
NOT_TIED = "не привязан"

HEADER_FILL = PatternFill("solid", fgColor="D9E2F3")
MANUAL_FILL = PatternFill("solid", fgColor="FFFF99")
TOTAL_FILL = PatternFill("solid", fgColor="EDEDED")
POOL_FILL = PatternFill("solid", fgColor="FCE4D6")
BOLD = Font(bold=True)
MONEY = "#,##0"
MONEY2 = "#,##0.00"
TX_ROWS = 5000  # сколько строк Транзакций покрыто выпадающими списками


def col(key):
    return get_column_letter(C[key] + 1)


def tx(key):
    """Колонка листа Транзакции целиком, например 'Транзакции'!$J:$J."""
    c = col(key)
    return f"'Транзакции'!${c}:${c}"


def read_csv(path):
    with open(path, newline="", encoding="utf-8") as f:
        rows = list(csv.reader(f))
    return rows[0], rows[1:]


def write_csv(path, header, rows):
    with open(path, "w", newline="", encoding="utf-8") as f:
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)


def batches():
    return read_csv(BATCH_CSV)[1]


def num(v):
    if v in ("", None):
        return None
    f = float(str(v).replace(" ", "").replace(",", "."))
    return int(f) if f.is_integer() else f


def amount(row):
    return num(row[C["amount"]]) or 0


def validate(row, batch_names):
    """Ошибки строки транзакции: такая строка не добавляется."""
    issues = []
    checks = [("cat", "categories", "категория"), ("art", "articles", "статья"),
              ("ptype", "product_types", "тип изделия"), ("season", "seasons", "сезон"),
              ("supplier", "suppliers", "поставщик"), ("alloc", "allocation_types", "тип отнесения")]
    for key, lst, label in checks:
        v = row[C[key]]
        if v and v not in LISTS[lst]:
            issues.append(f"{label} «{v}» не из справочника")
    if not row[C["cat"]]:
        issues.append("пустая категория")
    if not row[C["alloc"]]:
        issues.append("пустой тип отнесения")
    if row[C["batch"]] and row[C["batch"]] not in batch_names:
        issues.append(f"партия «{row[C['batch']]}» не из листа Партии")
    try:
        dt.datetime.strptime(row[C["date"]], "%d.%m.%Y")
    except ValueError:
        issues.append(f"дата «{row[C['date']]}» не в формате ДД.ММ.ГГГГ")
    try:
        if num(row[C["amount"]]) is None:
            issues.append("нет суммы")
    except ValueError:
        issues.append(f"сумма «{row[C['amount']]}» не число")
    return issues


def warnings(row):
    """Некритичные замечания: строка добавляется, но их стоит показать пользователю."""
    out = []
    if row[C["alloc"]] == "Прямой" and not row[C["batch"]]:
        out.append("прямой расход без партии — в себестоимость не попадет, в Сводке в строке «Прямые без партии»")
    if row[C["cat"]] == OTHER and not row[C["art"]]:
        out.append("прочий расход без статьи — в Сводке в строке «без статьи»")
    return out


def cmd_add(args):
    src = open(args[0], encoding="utf-8") if args else sys.stdin
    items = json.load(src)
    if isinstance(items, dict):
        items = [items]
    today = dt.datetime.now(ZoneInfo("Europe/Moscow")).strftime("%d.%m.%Y")
    header, rows = read_csv(TX_CSV)
    by_batch = {b[0]: b for b in batches()}
    new = []
    for it in items:
        row = [""] * len(TX_COLS)
        for k, v in it.items():
            if k not in JSON_KEYS:
                sys.exit(f"неизвестное поле {k!r}")
            row[JSON_KEYS[k]] = "" if v is None else str(v).strip()
        row[C["date"]] = row[C["date"]] or today
        row[C["amount"]] = str(num(row[C["amount"]]))
        if row[C["boxes"]]:
            row[C["boxes"]] = str(num(row[C["boxes"]]))
        b = by_batch.get(row[C["batch"]])
        if b:  # тип, сезон и поставщик берем из партии, если не заданы
            defaults = {"ptype": ("", b[1]), "season": (NOT_TIED, b[2]), "supplier": ("не указан", b[3])}
            for key, (empty, value) in defaults.items():
                if row[C[key]] in ("", empty):
                    row[C[key]] = value
        row[C["season"]] = row[C["season"]] or NOT_TIED
        row[C["supplier"]] = row[C["supplier"]] or "не указан"
        issues = validate(row, by_batch)
        if issues:
            sys.exit(f"строка не добавлена: {row}\n  " + "\n  ".join(issues))
        for w in warnings(row):
            print(f"внимание: {w}: {row[C['amount']]} | {row[C['comment']]}")
        new.append(row)
    write_csv(TX_CSV, header, rows + new)
    first = len(rows) + 2
    for i, r in enumerate(new):
        print(f"строка {first + i}: {r[C['date']]} | {r[C['cat']]} | {r[C['art']] or '—'} | "
              f"{r[C['batch']] or '—'} | {r[C['alloc']]} | {r[C['amount']]} | {r[C['comment']]}")


def cmd_batch(args):
    name, updates = args[0], args[1:]
    header, rows = read_csv(BATCH_CSV)
    for r in rows:
        if r[0] == name:
            for u in updates:
                k, v = u.split("=", 1)
                r[BATCH_KEYS[k]] = str(num(v))
            write_csv(BATCH_CSV, header, rows)
            print(r)
            return
    sys.exit(f"нет партии {name!r}; есть: {[r[0] for r in rows]}")


def cmd_check(_):
    _, rows = read_csv(TX_CSV)
    names = {b[0] for b in batches()}
    bad = 0
    for i, r in enumerate(rows, start=2):
        for issue in validate(r, names) + warnings(r):
            bad += 1
            print(f"строка {i}: {issue} | {r[C['amount']]} | {r[C['comment']]}")
    total = sum(amount(r) for r in rows)
    print(f"строк: {len(rows)}, сумма: {total:,.0f}, замечаний: {bad}".replace(",", " "))


def style_header(ws, row, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row, c)
        cell.font = BOLD
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center")


def style_row(ws, row, ncols, fill, bold=False):
    for c in range(1, ncols + 1):
        ws.cell(row, c).fill = fill
        if bold:
            ws.cell(row, c).font = BOLD


def widths(ws, values):
    for i, w in enumerate(values, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


def sheet_instructions(wb):
    ws = wb.active
    ws.title = "Инструкция"
    lines = [
        ("Учет расходов 2Flex2Sooul", "title"),
        ("Платежи пишите в чат Claude (текст, скриншоты, PDF) — он разносит их сюда. Таблица каждый раз выкладывается заново "
         "в папку «2Flex2Sooul — Учет расходов»: открывайте и расшаривайте папку. Правки лучше писать в чат.", None),
        ("Транзакции — строка на платеж. «Статья» — для прочих расходов. Красным — не попадет в себестоимость, оранжевым — прочие без статьи.", None),
        ("Партии — количества из приемки (желтые ячейки), без них нет себестоимости на единицу.", None),
        ("Сводка — все расходы: партии, общие и смешанные, прямые без партии; ИТОГО = все расходы. Ниже — сверка, общие по сезонам, "
         "категории с разбивкой прочих по статьям.", None),
        ("Прямые расходы идут в свою партию. Общие и смешанные с сезоном делятся между партиями сезона, с сезоном «не привязан» — "
         "между всеми партиями, пропорционально единицам.", None),
        ("Открытые вопросы: делить общие по единицам или коробкам; какая цифра на единицу основная; размерники и этикетки — "
         "сверить с Полиной.", None),
    ]
    for i, (text, kind) in enumerate(lines, start=2):
        c = ws.cell(i, 2, text)
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if kind == "title":
            c.font = Font(bold=True, size=14)
        elif kind == "h":
            c.font = Font(bold=True, size=12)
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 120


def sheet_transactions(wb, rows, nb):
    ws = wb.create_sheet("Транзакции")
    ncols = len(TX_COLS)
    ws.append(TX_COLS)
    style_header(ws, 1, ncols)
    for i, r in enumerate(rows, start=2):
        vals = list(r)
        vals[C["date"]] = dt.datetime.strptime(r[C["date"]], "%d.%m.%Y").date()
        vals[C["boxes"]] = num(r[C["boxes"]])
        vals[C["amount"]] = num(r[C["amount"]])
        ws.append([None if v == "" else v for v in vals])
        ws.cell(i, C["date"] + 1).number_format = "DD.MM.YYYY"
        ws.cell(i, C["amount"] + 1).number_format = MONEY
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(ncols)}{max(len(rows) + 1, 2)}"
    widths(ws, [11, 30, 20, 20, 13, 20, 30, 14, 11, 13, 60])
    # подсветка проблемных строк: красный — не попадет в себестоимость, оранжевый — прочие без статьи
    area = f"A2:{get_column_letter(ncols)}{TX_ROWS + 1}"
    L = {k: "$" + col(k) + "2" for k in C}
    red = f'OR(AND({L["amount"]}="",{L["cat"]}<>""),AND({L["cat"]}<>"",{L["alloc"]}=""),AND({L["alloc"]}="Прямой",{L["batch"]}=""))'
    ws.conditional_formatting.add(area, FormulaRule(formula=[red], fill=PatternFill("solid", bgColor="F4CCCC")))
    orange = f'AND({L["cat"]}="{OTHER}",{L["art"]}="")'
    ws.conditional_formatting.add(area, FormulaRule(formula=[orange], fill=PatternFill("solid", bgColor="FCE5CD")))
    lists = [("cat", "A", "categories"), ("art", "C", "articles"), ("ptype", "E", "product_types"),
             ("season", "G", "seasons"), ("supplier", "I", "suppliers"), ("alloc", "K", "allocation_types")]
    refs = [(key, f"'Справочники'!${c}$2:${c}${1 + len(LISTS[lst])}") for key, c, lst in lists]
    refs.append(("batch", f"'Партии'!$A$4:$A${3 + nb}"))
    for key, ref in refs:
        dv = DataValidation(type="list", formula1=ref, allow_blank=True)
        dv.add(f"{col(key)}2:{col(key)}{TX_ROWS + 1}")
        ws.add_data_validation(dv)


def sheet_batches(wb, brows):
    ws = wb.create_sheet("Партии")
    ws["A1"] = "Справочник партий — количества из приемки"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "Желтые колонки заполняются руками (или сообщением в чат). Без единиц в партии себестоимость на единицу не посчитается."
    ws.append(BATCH_COLS)
    style_header(ws, 3, len(BATCH_COLS))
    for r in brows:
        ws.append([None if v == "" else (num(v) if i >= 4 else v) for i, v in enumerate(r)])
        for c in range(5, 10):
            ws.cell(ws.max_row, c).fill = MANUAL_FILL
    ws.freeze_panes = "A4"
    widths(ws, [32, 22, 13, 22, 14, 12, 14, 14, 14])


def sheet_summary(wb, brows):
    ws = wb.create_sheet("Сводка")
    L = get_column_letter
    cats = LISTS["categories"]
    SUM, CAT, ART, BATCH, ALLOC, SEASON, DATE = (tx(k) for k in ("amount", "cat", "art", "batch", "alloc", "season", "date"))
    nb = len(brows)
    first, last = 4, 3 + nb
    pool_row, nobatch_row, total_row = last + 1, last + 2, last + 3
    cat_first, cat_last = 6, 5 + len(cats)
    c_sum, c_share, c_total, c_unit, c_ship = (cat_last + k for k in range(1, 6))
    ncols = c_ship

    ws["A1"] = "Сводка: все расходы, по партиям и себестоимость на единицу"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = "Все цифры считаются формулами из листов Транзакции и Партии. ИТОГО = все расходы."
    head = ["Артикул/партия", "Тип изделия", "Сезон", "Единиц в партии", "Отгружено единиц"] + cats + [
        "Итого по категориям", "Доля общих и смешанных", "Всего расходов", "На единицу (весь объем)", "На единицу (отгружено)"]
    for c, h in enumerate(head, start=1):
        ws.cell(3, c, h)
    style_header(ws, 3, ncols)
    ws.row_dimensions[3].height = 45

    units = f"$D${first}:$D${last}"
    # вспомогательный блок «Общие и смешанные по сезонам» — пул и единицы по каждому сезону, из него считается доля партий
    seasons_list = LISTS["seasons"]
    recon_row = total_row + 2
    pool_hdr = recon_row + 7
    p_first, p_last = pool_hdr + 2, pool_hdr + 1 + len(seasons_list)
    g = p_first + seasons_list.index(NOT_TIED)  # строка «не привязан» — пул на весь ассортимент
    lookup = f"$A${p_first}:$C${p_last}"
    for i, b in enumerate(brows):
        r = first + i
        ws.cell(r, 1, b[0])
        ws.cell(r, 2, b[1])
        ws.cell(r, 3, b[2])
        ws.cell(r, 4, f"=IFERROR(INDEX('Партии'!$E:$E,MATCH($A{r},'Партии'!$A:$A,0))*1,0)")
        ws.cell(r, 5, f"=IFERROR(INDEX('Партии'!$I:$I,MATCH($A{r},'Партии'!$A:$A,0))*1,0)")
        for c in range(cat_first, cat_last + 1):
            ws.cell(r, c, f'=SUMIFS({SUM},{BATCH},$A{r},{CAT},{L(c)}$3,{ALLOC},"Прямой")')
        ws.cell(r, c_share,
                f'=IF(OR($C{r}="{NOT_TIED}",$C{r}=""),0,IFERROR(VLOOKUP($C{r},{lookup},2,0)*$D{r}/VLOOKUP($C{r},{lookup},3,0),0))'
                f"+IFERROR($B${g}*$D{r}/$C${g},0)")
        ws.cell(r, c_unit, f'=IFERROR({L(c_total)}{r}/$D{r},"—")').number_format = MONEY2
        ws.cell(r, c_ship, f'=IFERROR({L(c_total)}{r}/$E{r},"—")').number_format = MONEY2

    ws.cell(pool_hdr, 1, "Общие и смешанные расходы по сезонам").font = Font(bold=True, size=12)
    for c, h in enumerate(["Сезон", "Общие и смешанные", "Единиц в партиях"], start=1):
        ws.cell(pool_hdr + 1, c, h)
    style_header(ws, pool_hdr + 1, 3)
    for k, season in enumerate(seasons_list, start=p_first):
        ws.cell(k, 1, season)
        crit = f'"{NOT_TIED}"' if season == NOT_TIED else f"$A{k}"
        pool = f'SUMIFS({SUM},{SEASON},{crit},{ALLOC},"Общий")+SUMIFS({SUM},{SEASON},{crit},{ALLOC},"Смешанный")'
        if season == NOT_TIED:  # пустой сезон тоже считаем «не привязан»; делится на все партии
            pool += f'+SUMIFS({SUM},{SEASON},"",{ALLOC},"Общий")+SUMIFS({SUM},{SEASON},"",{ALLOC},"Смешанный")'
            ws.cell(k, 3, f"=SUM({units})")
        else:
            ws.cell(k, 3, f"=SUMIF($C${first}:$C${last},$A{k},{units})")
        ws.cell(k, 2, "=" + pool).number_format = MONEY
        ws.cell(k, 3).number_format = MONEY
    ws.cell(g, 1, f"{NOT_TIED} (на все партии)")
    ws.cell(g, 1).font = BOLD

    # общие и смешанные: весь пул; в «Доле» с минусом то, что уже ушло в партии
    ws.cell(pool_row, 1, "Общие и смешанные расходы")
    ws.cell(pool_row, 2, "весь ассортимент")
    for c in range(cat_first, cat_last + 1):
        ws.cell(pool_row, c, f'=SUMIFS({SUM},{CAT},{L(c)}$3,{ALLOC},"Общий")+SUMIFS({SUM},{CAT},{L(c)}$3,{ALLOC},"Смешанный")')
    ws.cell(pool_row, c_share, f"=-SUM({L(c_share)}{first}:{L(c_share)}{last})")
    # прямые без партии и все, что не попало выше, — остаток, чтобы ИТОГО сходилось со всеми расходами
    ws.cell(nobatch_row, 1, "Прямые без партии")
    ws.cell(nobatch_row, 2, "в себестоимость не входят")
    for c in range(cat_first, cat_last + 1):
        ws.cell(nobatch_row, c, f"=SUMIFS({SUM},{CAT},{L(c)}$3)-SUM({L(c)}{first}:{L(c)}{pool_row})")
    ws.cell(nobatch_row, c_share, 0)
    for r in range(first, nobatch_row + 1):
        ws.cell(r, c_sum, f"=SUM({L(cat_first)}{r}:{L(cat_last)}{r})")
        ws.cell(r, c_total, f"={L(c_sum)}{r}+{L(c_share)}{r}")
        for c in range(4, c_total + 1):
            ws.cell(r, c).number_format = MONEY
    for r in (pool_row, nobatch_row):
        style_row(ws, r, ncols, POOL_FILL)

    ws.cell(total_row, 1, "ИТОГО — все расходы")
    for c in range(4, c_total + 1):
        ws.cell(total_row, c, f"=SUM({L(c)}{first}:{L(c)}{nobatch_row})").number_format = MONEY
    style_row(ws, total_row, ncols, TOTAL_FILL, bold=True)

    # сверка
    r = recon_row
    ws.cell(r, 1, "Сверка").font = Font(bold=True, size=12)
    tc = L(c_total)
    recon = [
        ("Всего расходов в Транзакциях", f"=SUM({SUM})"),
        ("Разнесено по партиям", f"=SUM({tc}{first}:{tc}{last})"),
        ("Не разнесено: общие и смешанные (нет единиц в Партиях или партий этого сезона)", f"={tc}{pool_row}"),
        ("Не разнесено: прямые без партии", f"={tc}{nobatch_row}"),
        ("Расхождение (строки с категорией не из списка) — должно быть 0", f"=B{r + 1}-{tc}{total_row}"),
    ]
    for k, (label, f) in enumerate(recon, start=1):
        ws.cell(r + k, 1, label)
        ws.cell(r + k, 2, f).number_format = MONEY
    ws.cell(r + 1, 1).font = BOLD
    ws.cell(r + 1, 2).font = BOLD

    # расходы по категориям, прочие — по статьям
    assert r + len(recon) + 2 == pool_hdr, "блок сезонов съехал"
    r = p_last + 2
    ws.cell(r, 1, "Расходы по категориям").font = Font(bold=True, size=12)
    for c, h in enumerate(["Категория / статья", "Сумма", "Доля, %", "Платежей", "Последняя оплата"], start=1):
        ws.cell(r + 1, c, h)
    style_header(ws, r + 1, 5)
    k = r + 2
    cat_rows = []

    def line(label, crit, indent=False):
        ws.cell(k, 1, ("      " if indent else "") + label)
        ws.cell(k, 2, f"=SUMIFS({SUM},{crit})").number_format = MONEY
        ws.cell(k, 3, f"=IFERROR(B{k}/$B${r + 2 + len(cats) + len(LISTS['articles']) + 1},0)").number_format = "0.0%"
        ws.cell(k, 4, f"=COUNTIFS({crit})")
        ws.cell(k, 5, f'=IF(D{k}=0,"",MAXIFS({DATE},{crit}))').number_format = "DD.MM.YYYY"
        if indent:
            for c in range(1, 6):
                ws.cell(k, c).font = Font(italic=True, color="595959")

    for cat in cats:
        cat_rows.append(k)
        line(cat, f'{CAT},"{cat}"')
        k += 1
        if cat == OTHER:
            for art in LISTS["articles"] + [""]:
                line(art or "без статьи", f'{CAT},"{OTHER}",{ART},"{art}"', indent=True)
                k += 1
    ws.cell(k, 1, "ИТОГО")
    ws.cell(k, 2, "=" + "+".join(f"B{x}" for x in cat_rows)).number_format = MONEY
    ws.cell(k, 3, f"=IFERROR(B{k}/B{k},0)").number_format = "0.0%"
    ws.cell(k, 4, "=" + "+".join(f"D{x}" for x in cat_rows))
    ws.cell(k, 5, f"=MAX(E{r + 2}:E{k - 1})").number_format = "DD.MM.YYYY"
    style_row(ws, k, 5, TOTAL_FILL, bold=True)
    assert k == r + 2 + len(cats) + len(LISTS["articles"]) + 1, "ссылка на ИТОГО в колонке «Доля» съехала"

    ws.freeze_panes = "B4"
    widths(ws, [44, 20, 13, 11, 16] + [13] * len(cats) + [13, 13, 13, 12, 12])


def sheet_lists(wb):
    ws = wb.create_sheet("Справочники")
    cols = [(1, "Категории расхода", "categories"), (3, "Статьи прочих расходов", "articles"),
            (5, "Типы изделий", "product_types"), (7, "Сезоны", "seasons"),
            (9, "Поставщик/производство", "suppliers"), (11, "Тип отнесения", "allocation_types")]
    for c, title, key in cols:
        ws.cell(1, c, title)
        ws.cell(1, c).font = BOLD
        ws.cell(1, c).fill = HEADER_FILL
        for i, v in enumerate(LISTS[key], start=2):
            ws.cell(i, c, v)
        ws.column_dimensions[get_column_letter(c)].width = 38 if c == 1 else 26
        ws.column_dimensions[get_column_letter(c + 1)].width = 3


def cmd_build(_):
    _, rows = read_csv(TX_CSV)
    brows = batches()
    wb = Workbook()
    sheet_instructions(wb)
    sheet_transactions(wb, rows, len(brows))
    sheet_batches(wb, brows)
    sheet_summary(wb, brows)
    sheet_lists(wb)
    wb.active = 1
    OUT.parent.mkdir(exist_ok=True)
    wb.save(OUT)
    print(f"{OUT} ({OUT.stat().st_size} байт, строк: {len(rows)})")


def cmd_summary(_):
    """Те же расчеты, что формулы листа Сводка, — для сверки с Google Таблицей и ответов в чате."""
    _, rows = read_csv(TX_CSV)
    brows = batches()
    units = {b[0]: num(b[4]) or 0 for b in brows}
    shipped = {b[0]: num(b[8]) or 0 for b in brows}
    season_of = {b[0]: b[2] for b in brows}
    total_units = sum(units.values())
    direct, pool = {}, {}
    pool_total = 0
    for r in rows:
        if r[C["alloc"]] == "Прямой" and r[C["batch"]] in units:
            direct[r[C["batch"]]] = direct.get(r[C["batch"]], 0) + amount(r)
        elif r[C["alloc"]] in ("Общий", "Смешанный"):
            s = r[C["season"]]
            key = s if s not in ("", NOT_TIED) else None
            pool[key] = pool.get(key, 0) + amount(r)
            pool_total += amount(r)

    def fmt(v):
        return f"{v:,.0f}".replace(",", " ")

    placed = 0
    for name in units:
        s = season_of[name]
        share = 0
        if s not in ("", NOT_TIED):
            season_units = sum(u for n, u in units.items() if season_of[n] == s)
            share += pool.get(s, 0) * units[name] / season_units if season_units else 0
        share += pool.get(None, 0) * units[name] / total_units if total_units else 0
        total = direct.get(name, 0) + share
        placed += total
        per = [f"{total / q:,.2f}".replace(",", " ") if q else "—" for q in (units[name], shipped[name])]
        print(f"{name}: прямые {fmt(direct.get(name, 0))}, доля общих {fmt(share)}, всего {fmt(total)}, "
              f"на ед. {per[0]}, на отгруж. {per[1]}")
    grand = sum(amount(r) for r in rows)
    placed_share = placed - sum(direct.values())
    print(f"Общие и смешанные: {fmt(pool_total)}, распределено {fmt(placed_share)}, осталось {fmt(pool_total - placed_share)}")
    print(f"Прямые без партии: {fmt(grand - sum(direct.values()) - pool_total)}")
    print(f"ИТОГО {fmt(grand)}, разнесено по партиям {fmt(placed)}, не разнесено {fmt(grand - placed)}")
    for cat in LISTS["categories"]:
        in_cat = [r for r in rows if r[C["cat"]] == cat]
        print(f"  {cat}: {fmt(sum(amount(r) for r in in_cat))} ({len(in_cat)})")
        if cat == OTHER:
            for art in LISTS["articles"] + [""]:
                in_art = [r for r in in_cat if r[C["art"]] == art]
                if in_art:
                    print(f"      {art or 'без статьи'}: {fmt(sum(amount(r) for r in in_art))} ({len(in_art)})")


def cmd_import(args):
    from openpyxl import load_workbook
    wb = load_workbook(args[0], data_only=True)

    def cell(v):
        if v is None:
            return ""
        if isinstance(v, (dt.datetime, dt.date)):
            return v.strftime("%d.%m.%Y")
        if isinstance(v, float):
            return str(num(v))
        return str(v).strip()

    ws = wb["Транзакции"]
    header = [cell(v) for v in next(ws.iter_rows(max_row=1, values_only=True))]
    pos = {h: i for i, h in enumerate(header)}  # колонки ищем по заголовку: старые выгрузки без «Статьи» тоже читаются
    tx_rows = []
    for r in ws.iter_rows(min_row=2, values_only=True):
        vals = [cell(r[pos[h]]) if h in pos and pos[h] < len(r) else "" for h in TX_COLS]
        if any(vals):
            tx_rows.append(vals)
    bt = [[cell(v) for v in r[:len(BATCH_COLS)]]
          for r in wb["Партии"].iter_rows(min_row=4, values_only=True) if r[0]]
    write_csv(TX_CSV, TX_COLS, tx_rows)
    write_csv(BATCH_CSV, BATCH_COLS, bt)
    print(f"Транзакций: {len(tx_rows)}, партий: {len(bt)}")
    cmd_check([])


if __name__ == "__main__":
    cmds = {"add": cmd_add, "batch": cmd_batch, "check": cmd_check, "build": cmd_build,
            "summary": cmd_summary, "import": cmd_import}
    if len(sys.argv) < 2 or sys.argv[1] not in cmds:
        sys.exit(__doc__)
    cmds[sys.argv[1]](sys.argv[2:])
