"""Recompute answers.json from the raw AK Furs exports.

Usage: python solve.py data/ > answers.json
Answers go to stdout; sanity counts go to stderr.
"""
import csv
import datetime as dt
import glob
import json
import os
import re
import sys
from collections import Counter

MONTHS = {'янв': 1, 'фев': 2, 'мар': 3, 'апр': 4, 'мая': 5, 'май': 5, 'июн': 6,
          'июл': 7, 'авг': 8, 'сен': 9, 'окт': 10, 'ноя': 11, 'дек': 12}


def log(*a):
    print(*a, file=sys.stderr)


def find(data_dir, prefix):
    hits = sorted(glob.glob(os.path.join(data_dir, prefix + '*')))
    if not hits:
        sys.exit(f'no file {prefix}* in {data_dir}')
    return hits[0]


def read_text(path):
    raw = open(path, 'rb').read()
    try:
        return raw.decode('utf-8-sig')
    except UnicodeDecodeError:
        return raw.decode('cp1251')  # bank export is Windows-1251


def order_no(s):
    """A-00075, А-75 (Cyrillic), a-75, A75 -> 75."""
    m = re.search(r'\d+', s)
    return int(m.group()) if m else None


def money(s):
    """'2,48 млн', '245,000.00', '140000,00', '2 200 000', '74.000' -> float."""
    s = s.strip().replace('\xa0', ' ').replace('₽', '').replace('руб.', '').replace('руб', '').strip()
    mult = 1
    if 'млн' in s:
        mult, s = 1_000_000, s.replace('млн', '').strip()
    s = s.replace(' ', '')
    if re.fullmatch(r'-?\d{1,3}(,\d{3})+(\.\d+)?', s):      # 245,000.00
        s = s.replace(',', '')
    elif re.fullmatch(r'-?\d{1,3}(\.\d{3})+', s):           # 74.000 (thousands dot)
        s = s.replace('.', '')
    else:                                                   # 140000,00 / 2,48
        s = s.replace(',', '.')
    return float(s) * mult


def parse_date(s):
    s = s.strip()
    if re.fullmatch(r'\d{5}', s):                           # Excel serial
        return dt.date(1899, 12, 30) + dt.timedelta(days=int(s))
    if re.fullmatch(r'\d{4}-\d{2}-\d{2}', s):
        return dt.date.fromisoformat(s)
    m = re.fullmatch(r'(\d{1,2})[./](\d{1,2})[./](\d{4})', s)  # day first: day values go up to 31
    if m:
        return dt.date(int(m[3]), int(m[2]), int(m[1]))
    m = re.fullmatch(r'(\d{1,2})\s+([а-я]+)\.?\s+(\d{4})', s.lower())
    if m and m[2][:3] in MONTHS:
        return dt.date(int(m[3]), MONTHS[m[2][:3]], int(m[1]))
    raise ValueError(f'unknown date format: {s!r}')


def status(s):
    s = s.strip().lower().replace('ё', 'е')
    if s in ('оплачен', 'оплачено', 'paid'):
        return 'Оплачен'
    if s in ('отменен', 'отменено', 'cancelled', 'canceled'):
        return 'Отменён'
    if s in ('возврат', 'refund'):
        return 'Возврат'
    return 'В работе'


def load_lines(path):
    """Dedup on (order number, line number), keeping the latest 'Выгружено'."""
    rows = list(csv.DictReader(read_text(path).splitlines(), delimiter=';'))
    latest = {}
    for r in rows:
        key = (order_no(r['Номер заказа']), int(r['№ строки']))
        if key not in latest or r['Выгружено'].strip() > latest[key]['Выгружено'].strip():
            latest[key] = r
    lines = []
    for (order, line), r in latest.items():
        lines.append({
            'order': order, 'line': line,
            'created': parse_date(r['Дата создания']),
            'manager_raw': r['Менеджер'],
            'sku': r['Артикул'].strip(),
            'qty': int(r['Кол-во']),
            'price': money(r['Цена за ед.']),
            'discount_raw': r['Скидка'].strip(),
            'status': status(r['Статус']),
        })
    return rows, lines


# Order reference in a payment purpose: "A877", "А-00391", "зак. 528", "заказ 775", "№ 00991".
# The letter must not follow another letter, so "Оплата 75" is not read as order 75.
ORDER_REF = re.compile(r'(?<![^\W\d_])(?:[AАaа]-?|№\s*|зак(?:аз[а-я]*)?\.?\s*)0*(\d+)', re.I)


def load_bank(path):
    rows = list(csv.reader(read_text(path).splitlines(), delimiter=';'))
    head = next(i for i, r in enumerate(rows) if r and r[0].strip() == 'Дата операции')
    cols = [c.strip() for c in rows[head]]
    return [dict(zip(cols, r)) for r in rows[head + 1:] if any(c.strip() for c in r)]


def write_check(name, header, rows):
    """Full lists for checking by eye go to checks/, the console gets a summary."""
    os.makedirs('checks', exist_ok=True)
    with open(os.path.join('checks', name), 'w', encoding='utf-8-sig', newline='') as f:
        w = csv.writer(f, delimiter=';')
        w.writerow(header)
        w.writerows(rows)
    return os.path.join('checks', name)


def rub(x):
    x = round(x, 2)
    return int(x) if x == int(x) else x


def q2(lines, bank, notes):
    orders = {l['order'] for l in lines}
    statuses = {}
    for l in lines:
        statuses.setdefault(l['order'], []).append(l['status'])
    paid_any = {o for o, st in statuses.items() if 'Оплачен' in st}  # any non-cancelled line Оплачен
    paid_all = {o for o, st in statuses.items() if set(st) == {'Оплачен'}}
    mixed = sum(1 for st in statuses.values() if len(set(st)) > 1)
    log(f'Q2: orders with mixed line statuses {mixed}; paid orders: any-line rule {len(paid_any)}, '
        f'all-lines rule {len(paid_all)}, differ by {len(paid_any ^ paid_all)}')

    incoming = [b for b in bank if b['Тип'].strip() == 'Поступление']
    log(f'Q2: bank rows {len(bank)}, incoming {len(incoming)}, skipped outgoing {len(bank) - len(incoming)}')
    paid_orders, unident, unread = set(), [], []
    for b in incoming:
        purpose = b['Назначение платежа']
        m = ORDER_REF.search(purpose)
        if m and int(m.group(1)) in orders:
            paid_orders.add(int(m.group(1)))
        else:
            unident.append(b)
        rest = purpose[:m.start()] + ' ' + purpose[m.end():] if m else purpose
        other = re.findall(r'[+\d][\d\s().:-]*\d', rest)
        if other:
            unread.append([b['Дата операции'], b['Сумма'], purpose, m.group(1) if m else '', ' | '.join(o.strip() for o in other)])

    def answer(paid):
        return sorted(o for o in paid if o not in paid_orders)

    q2a, q2a_all = answer(paid_any), answer(paid_all)
    total = sum(money(b['Сумма']) for b in unident)
    log(f'Q2a: paid orders without any incoming payment: {len(q2a)} -> {q2a}')
    if q2a != q2a_all:
        log(f'Q2a: with the all-lines rule the answer would be {len(q2a_all)} -> {q2a_all}')
        notes.append(f'Q2a: заказ считаю «Оплачен», если хотя бы одна неотменённая строка «Оплачен». '
                     f'Если требовать «Оплачен» у всех строк, ответ другой: {q2a_all}.')
    else:
        notes.append(f'Q2a: заказ считаю «Оплачен», если хотя бы одна неотменённая строка «Оплачен» '
                     f'(правило «все строки Оплачен» даёт тот же ответ: смешанных статусов в заказах {mixed}).')
    log(f'Q2b: unidentified incoming payments {len(unident)}, sum {rub(total)}')
    for b in unident:
        log(f'   {b["Дата операции"]}  {b["Сумма"]:>14}  {b["Назначение платежа"]}')
    by_shape = Counter(re.sub(r'\d+', '9', r[2]) for r in unread)
    path = write_check('q2_unread_numbers.csv', ['дата', 'сумма', 'назначение', 'прочитан номер заказа', 'прочие числа'], unread)
    log(f'Q2: {len(unread)} purposes contain numbers not read as an order number (full list: {path}):')
    for shape, n in by_shape.most_common():
        log(f'   {n:4}  {shape}')
    notes.append('Q2: платёж опознан, если в назначении есть номер существующего заказа (в т.ч. отменённого): '
                 '«A877», «А-00391», «зак. 528», «заказ 775», а также «Счёт №/по счету №N» — номер счёта считаю '
                 'номером заказа (суммы таких платежей совпадают с суммами заказов так же часто, как у «Оплата заказа»). '
                 'Телефоны и даты в назначении игнорируются. Списания (возвраты) не считаются поступлениями.')
    return {
        'q2a_paid_orders_without_payment': q2a,
        'q2b_unidentified_payments_count': len(unident),
        'q2b_unidentified_payments_sum_rub': rub(total),
    }


def main(data_dir):
    notes = []
    raw_rows, lines = load_lines(find(data_dir, '01_'))
    orders = {l['order'] for l in lines}
    log(f'Q1: raw rows {len(raw_rows)}, unique lines {len(lines)}, '
        f'dropped duplicates {len(raw_rows) - len(lines)}, unique orders {len(orders)} '
        f'(numbers {min(orders)}..{max(orders)})')
    log('Q1: statuses per line', dict(Counter(l['status'] for l in lines)))
    notes.append('Q1: номер заказа — цифры из поля; дубли по (заказ, № строки) — берётся строка с последним «Выгружено». '
                 'Даты вида 11/02/2025 читаю как ДД/ММ (первое число доходит до 31). '
                 'Статусы: NEW, «Ждём оплату» и прочее -> «В работе»; статус берётся только из поля «Статус», не из комментария.')

    answers = json.load(open(find(data_dir, 'answers_template')))
    answers.update({
        'q1_unique_orders': len(orders),
        'q1_unique_order_lines': len(lines),
    })
    answers.update(q2(lines, load_bank(find(data_dir, '02_')), notes))
    answers['run_command'] = 'python solve.py data/ > answers.json'
    answers['notes'] = ' '.join(notes)
    json.dump(answers, sys.stdout, ensure_ascii=False, indent=1)
    print()


if __name__ == '__main__':
    main(sys.argv[1] if len(sys.argv) > 1 else 'data/')
